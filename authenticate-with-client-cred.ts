import * as dotenv from "dotenv";
import axios from "axios";
import Express from "express";
import open from "open";

dotenv.config();
const port = 3000;
const REDIRECT_URI = `http://localhost:${port}/callback`;
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

class Authenticator {
  constructor(private port: number) {}

  async authenticate(): Promise<string> {
    const authCodeUrl = await this.getAuthCodeUrl();
    console.log("Authenticate in your browser through this url: ", authCodeUrl);
    open(authCodeUrl);
    const authCode = await this.waitForCode();
    const token = await this.exchangeCode(authCode);
    return token;
  }

  private async getAuthCodeUrl(): Promise<string> {
    const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(
      REDIRECT_URI
    )}&response_mode=query&scope=${scopes}`;
    return url;
  }

  private waitForCode(): Promise<string> {
    const app = Express();

    return new Promise((resolve) => {
      let calledBack = false;
      app.get("/callback", (req: Express.Request, res) => {
        calledBack = true;
        const code = req.query.code;
        res.send(`<html><body><script>window.close()</script></body></html>`);
        resolve(code as string);
        server.close();
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

  private async exchangeCode(code: string): Promise<string> {
    try {
      const params = new URLSearchParams({
        client_id: CLIENT_ID || "",
        client_secret: CLIENT_SECRET || "",
        code: code,
        redirect_uri: REDIRECT_URI,
        grant_type: "client_credentials",
        scope: `00000003-0000-0ff1-ce00-000000000000/.default`,
      });
      const response = await axios.post(
        `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
        params,
        {
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
        }
      );
      const accessToken = response.data.access_token;
      console.log("Access Token:", accessToken);
      return accessToken;
    } catch (error) {
      console.error("Error getting access token:", error);
    }
    return "";
  }
}

const authenticator = new Authenticator(port);
authenticator.authenticate().catch(console.error);
