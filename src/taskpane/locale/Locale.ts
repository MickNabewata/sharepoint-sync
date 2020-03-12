import { LocaleEn } from "./strings/en-us";
import { LocaleJp } from "./strings/ja-jp";

/** 文言定義 */
export interface ILocale {
    /** 認証画面用 */
    Auth: {
        /** セクションタイトル */
        SectionTitle: string;
        /** ログイン依頼 */
        PleaseLogin: string;
        /** ドメイン入力テキストボックスのプレースホルダ */
        DomainPlaceHolder: string;
        /** ログインボタン */
        LoginButton: string;
        /** ログアウトボタン */
        LogoutButton: string;
    }
}

/** Office環境に応じた文言定義を取得 */
export function GetLocale(): ILocale {
    switch (Office.context.displayLanguage) {
        case "ja-JP":
            return LocaleJp;
        case "en-US":
            return LocaleEn;
        default:
            return LocaleEn;
    }
}
