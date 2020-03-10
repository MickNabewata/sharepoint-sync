import * as React from "react";
import { AuthError, Account } from "msal";
import { PnPFetchClient } from "../pnp/PnPFetchClient";
import "@pnp/polyfill-ie11";
import { sp } from '@pnp/sp';
import { msalInstance } from "../pnp/MSal";
import { Spinner, SpinnerType } from "office-ui-fabric-react";
import QueryUtil from "../util/queryUtil";
import { stringIsNullOrEmpty } from "@pnp/common";

/** ログインページプロパティ */
export interface MsalRedirectProps {
}

/** ログインページステート */
export interface MsalRedirectStates {
  /** ドメイン */
  domain: string;
  /** 認証済か否か */
  authenticated: boolean;
  /** エラーメッセージ */
  errorMessage: string;
}

/** ダイアログメッセージ */
export interface DialogMessage {
  /** メッセージ */
  message: string;
  /** アカウント */
  account: Account;
  /** SharePointアクセストークン */
  token: string;
}

let pnpFetchClient: PnPFetchClient;

/** ログインページ */
export default class MsalRedirect extends React.Component<MsalRedirectProps, MsalRedirectStates> {

  /** ログインページ  */
  constructor(props, context) {
    super(props, context);

    // ステートを初期化
    this.state = {
      domain: this.getDomain(),
      authenticated: false,
      errorMessage: ''
    };

    // MSALインスタンスを初期化
    this.initMsal();
  }

  /** URLパラメータからドメインを取得 */
  private getDomain() {
    let ret = '';
    const params = new QueryUtil().get().params;
    if (params && !stringIsNullOrEmpty(params.domain)) {
      ret = params.domain;
    }
    return ret;
  }

  /** MSALインスタンスを初期化 */
  private initMsal() {
    // 認証コールバックを仕掛ける
    msalInstance.handleRedirectCallback(
      async () => {
        // 成功を記録してPnPを初期化
        await this.setToState({ authenticated: true });
        this.initPnPjs();
      },
      async (authErr: AuthError, accountState: string) => {
        // 失敗を記録
        console.log(authErr);
        console.log(accountState);
        await this.setToState({ errorMessage: authErr.errorMessage });
      }
    );
  }

  /** PnPを初期化 */
  private initPnPjs(): void {
    const { domain } = this.state;

    pnpFetchClient = new PnPFetchClient(msalInstance);

    const fetchClientFactory = () => {
      return pnpFetchClient;
    };

    sp.setup({
      sp: {
        fetchClientFactory,
        baseUrl: `https://${domain}.sharepoint.com/`
      }
    });
  }

  /** ステートにセット */
  private setToState(state: Partial<MsalRedirectStates>): Promise<void> {
    return new Promise<void>((resolve: () => void) => {
      this.setState(state as MsalRedirectStates, () => {
        resolve();
      });
    });
  }

  /** レンダリング */
  public render() {
    const { authenticated } = this.state;
    return (
      <div>
        {authenticated ? undefined : <Spinner type={SpinnerType.large} />}
        <div>{this.state.errorMessage}</div>
      </div>
    );
  }

  /** コンポーネントがマウントされた後のイベント */
  public componentDidMount(): void {
    // URLハッシュにアクセストークンが含まれている（コールバック時）場合はここで終了
    if (msalInstance.isCallback(window.location.hash) === false) {
      // ログイン処理
      const account = msalInstance.getAccount();
      if (!account) {
        // 未認証なのでログインページにリダイレクトする
        msalInstance.loginRedirect({});
      } else {
        // 認証済
        this.setToState({ authenticated: true }).then(() => {
          
          // PnPを初期化し、一度問合せをしてアクセストークンを得る
          this.initPnPjs();
          try {
            sp.web.get().then(
              () => {
                const message: DialogMessage = {
                  message: "Success",
                  account: account,
                  token: pnpFetchClient.token
                };
                // ダイアログを閉じる
                Office.context.ui.messageParent(JSON.stringify(message));
              },
              (err) => {
                console.log(err);

                // ダイアログを閉じる
                const message: DialogMessage = {
                  message: "login failed. please check your domain and account.",
                  account: account,
                  token: pnpFetchClient.token
                };
                Office.context.ui.messageParent(JSON.stringify(message));

                //this.setToState({ errorMessage: err ? err.description ? err.description : err.error_description ? err.error_description : JSON.stringify(err) : "error" });
              }
            );
          } catch (ex) {
            console.log(ex);

            // ダイアログを閉じる
            const message: DialogMessage = {
              message: "login failed. please check your domain and account.",
              account: account,
              token: pnpFetchClient.token
            };
            Office.context.ui.messageParent(JSON.stringify(message));
            
            //this.setToState({ errorMessage: ex ? JSON.stringify(ex) : "exception" });
          }
        });
      }
    }
  }
}