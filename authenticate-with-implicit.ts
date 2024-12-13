import * as dotenv from "dotenv";
import axios from "axios";
import Express from "express";
import open from "open";

dotenv.config();
const port = 3000;
const REDIRECT_URI = `http://localhost:${port}/callback?`;
const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
const scopes = [
  "MyFiles.Read",
  "MyFiles.Write",
  "Container.Selected",
  "Sites.FullControl.All",
  "Sites.Manage.All",
  "User.ReadWrite.All",
]
  .map(encodeURIComponent)
  .join("%20");
const audience = `00000003-0000-0ff1-ce00-000000000000/.default`;

class Authenticator {
  constructor(private port: number) {}

  async authenticate(): Promise<string> {
    const authCodeUrl = await this.getAuthUrl();
    console.log("Authenticate in your browser through this url: ", authCodeUrl);
    open(authCodeUrl);
    const token = await this.waitForToken();
    return token;
  }

  private async getAuthUrl(): Promise<string> {
    const url =
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
      `client_id=${CLIENT_ID}` +
      `&response_type=token` +
      `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
      `&scope=${audience}` +
      `&state=SomeState` +
      `&nonce=dosfindoisfns`;
    return url;
  }

  private waitForToken(): Promise<string> {
    const app = Express();

    return new Promise((resolve) => {
      let calledBack = false;
      app.get("/callback", (req: Express.Request, res) => {
        // This is a hack to get the token from the URL hash. We load the token from the page,
        // then do a second callback to the server to get the token.
        const code = req.query.access_token;
        res.send(
          `<html><body><div id="token" /><script>
          async function fetchToken() {
            const tokenElement = document.getElementById("token");
            const hash = window.location.hash.substring(1);
            const elements = hash.split("&");
            const dict = elements.reduce((acc, element) => {
              const [key, value] = element.split("=");
              acc[key] = value;
              return acc;
            }, {});
            tokenElement.innerText = "Token: " + dict.access_token;
            await fetch("/token?token=" + dict.access_token);
            window.close();
          };
          fetchToken().then();
          </script></body></html>`
        );
        res.end();
      });

      app.get("/token", (req: Express.Request, res) => {
        calledBack = true;
        const token = req.query.token;
        res.end();
        server.close();
        resolve(token as string);
      });

      setTimeout(() => {
        if (!calledBack) {
          server.close();
          throw new Error("Timeout");
        }
      }, 10000);

      const server = app.listen(port, () => {
        console.log(`Listening on port ${port}`);
      });
    });
  }
}

const authenticator = new Authenticator(port);
authenticator
  .authenticate()
  .then((token) => {
    console.log(token);
  })
  .catch(console.error);
