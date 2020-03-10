import { UserAgentApplication } from "msal";
import { AuthOptions } from "msal/lib-commonjs/Configuration";

/** MSAL認証用 */
const msalConfig: AuthOptions = {
  authority: "https://login.microsoftonline.com/9e129fc5-a933-4f8f-a17b-f71057e7d820",
  clientId: "f8208612-c057-4d8e-b987-348ed1d45bc9",
  redirectUri: "https://sharepoint-importer.web.app/taskpane.html#/login"
};

/** MSAL認証インスタンス */
export const msalInstance: UserAgentApplication = new UserAgentApplication({
  auth: msalConfig
});