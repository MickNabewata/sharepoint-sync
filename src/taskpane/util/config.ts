/** 環境変数取得ユーティリティ */
export default class Config {
    /** アドインホストURL */
    public static get HOST() { return process.env.REACT_APP_HOST }
}