import { UserAgentApplication } from "msal";
import { AuthOptions } from "msal/lib-commonjs/Configuration";
import Config from "../util/config";

/** MSAL認証用 */
const msalConfig: AuthOptions = {
  authority: "https://login.microsoftonline.com/organizations",
  clientId: Config.CLIENT_ID,
  redirectUri: `${Config.HOST}/taskpane.html#/login`
};

/** MSAL認証インスタンス */
export const msalInstance: UserAgentApplication = new UserAgentApplication({
  auth: msalConfig
});