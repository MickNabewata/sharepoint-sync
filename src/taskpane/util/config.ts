/** 環境変数取得ユーティリティ */
export default class Config {
    /** アドインホストURL */
    public static get HOST() { return process.env.REACT_APP_HOST }
    /** 認証先Azure ADアプリケーションのクライアントID */
    public static get CLIENT_ID() { return process.env.REACT_APP_CLIENT_ID }
}