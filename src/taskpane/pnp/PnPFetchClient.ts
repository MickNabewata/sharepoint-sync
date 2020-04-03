import { BearerTokenFetchClient, FetchOptions, isUrlAbsolute } from '@pnp/common';
import { UserAgentApplication, AuthenticationParameters } from 'msal';

/** PnP認証クライアント */
export class PnPFetchClient extends BearerTokenFetchClient {
  constructor(private authContext: UserAgentApplication) {
    super(null);
  }

  /** フェッチ */
  public async fetch(url: string, options: FetchOptions = {}): Promise<Response> {
    if (!isUrlAbsolute(url)) {
      throw new Error('You must supply absolute urls to PnPFetchClient.fetch.');
    }

    const token = await this.getToken(this.getResource(url));
    this.token = token;
    return super.fetch(url, options);
  }

  /** アクセストークン取得 */
  private async getToken(resource: string): Promise<string> {
    return new Promise<string>(async (resolve: (val: string) => void, reject: (err?) => void) => {
      const request: AuthenticationParameters = {
      };

      if (resource.indexOf('sharepoint') !== -1) {
        request.scopes = [`${resource}/AllSites.Write`];
      }

      try {
        const response = await this.authContext.acquireTokenSilent(request);
        resolve(response.accessToken);
      } catch (error) {
        if (this.requiresInteraction(error.errorCode)) {
          this.authContext.acquireTokenRedirect(request);
          reject();
        } else {
          reject(error);
        }
      }
    });
  }

  /** ユーザーと対話する必要の有無 */
  private requiresInteraction(errorCode: string) {
    if (!errorCode || !errorCode.length) {
      return false;
    }
    return errorCode === "consent_required" ||
      errorCode === "interaction_required" ||
      errorCode === "login_required";
  }

  /** リソースURL取得 */
  private getResource(url: string): string {
    const parser = document.createElement('a') as HTMLAnchorElement;
    parser.href = url;
    return `${parser.protocol}//${parser.hostname}`;
  }
}
