import * as React from "react";
import Section from "./Section";
import { Text, PrimaryButton, TextField } from "office-ui-fabric-react";
import { DialogMessage } from "./MsalRedirect";
import { Account } from "msal";
import { stringIsNullOrEmpty } from "@pnp/common";
import Config from "../util/config";

/** 認証プロパティ */
export interface AuthProps {
  /** 認証成功コールバック */
  authCallBack: (account: Account, token: string, domain: string) => void;
  /** ログアウトコールバック */
  logOutCallBack: () => void;
}

/** 認証ステート */
export interface AuthStates {
  /** ドメイン */
  domain: string;
  /** エラーメッセージ */
  errorMessage: string;
}

/** 認証 */
export default class Auth extends React.Component<AuthProps, AuthStates> {

  /** 認証  */
  constructor(props, context) {
      super(props, context);
    this.state = {
        domain: '',
        errorMessage: ''
      };
  }

  /** 認証ダイアログクローズ時コールバック */
  private authCallback = (dialog: Office.Dialog) => (args: any) => {
    const { domain } = this.state;
    const { authCallBack } = this.props;
    dialog.close();
    const dialogMessage: DialogMessage = JSON.parse(args.message);
    if (dialogMessage && dialogMessage.message === "Success") {
      if (authCallBack) authCallBack(dialogMessage.account, dialogMessage.token, domain);
    } else {
      this.setState({ errorMessage: dialogMessage ? dialogMessage.message : args ? JSON.stringify(args) : "args is null." });
    }
  }

  /** ドメイン入力イベント */
  private handleDomainChanged = (event: React.ChangeEvent<HTMLInputElement>, newValue: string) => {
    console.log(event);
    console.log(newValue);

    this.setState({ domain: newValue });
  }

  /** ログインボタンクリックイベント */
  private handleLoginButtonClicked = () => {
    const { domain } = this.state;

    Office.context.ui.displayDialogAsync(`${Config.HOST}/taskpane.html?domain=${domain}#/login`, {}, result => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, this.authCallback(dialog));
    });
  }

  /** ログアウトボタンクリックイベント */
  private handleLogOutButtonClicked = () => {
    const { logOutCallBack: singOutCallBack } = this.props;
    if (singOutCallBack) singOutCallBack();
  }

  /** レンダリング */
  public render() {
    const { domain, errorMessage } = this.state;

    return (
      <Section title="認証が必要です">
        <Text>取込先のSharePointリストに投稿以上の権限を持つアカウントでログインしてください。</Text>
        <div className="ex-sp__section-item">
          <TextField
            label="https://"
            suffix=".sharepoint.com"
            required
            underlined
            placeholder="ドメイン"
            value={domain}
            onChange={this.handleDomainChanged}
          />
        </div>
        <div className="ex-sp__section-item">
          <PrimaryButton
            text="ログイン"
            onClick={this.handleLoginButtonClicked}
            disabled={stringIsNullOrEmpty(domain)}
          />
          <PrimaryButton
            text="ログアウト"
            onClick={this.handleLogOutButtonClicked}
            className="ex-sp__section-column"
          />
        </div>
        <div className="ex-sp__section-item">
          <Text>{errorMessage}</Text>
        </div>
      </Section>
    );
  }
}