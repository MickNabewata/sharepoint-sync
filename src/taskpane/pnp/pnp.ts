import { SPRest } from "@pnp/sp";
import "@pnp/polyfill-ie11";
import { stringIsNullOrEmpty } from "@pnp/common";

/** PnP初期化 */
export function initPnPJs(sp: SPRest, token: string, selectedWebUrl: string) {
    if (stringIsNullOrEmpty(token) || stringIsNullOrEmpty(selectedWebUrl)) return;
    sp.setup({
        sp: {
            headers: {
                "Authorization": `Bearer ${token}`
            },
            baseUrl: selectedWebUrl
        }
    });
}