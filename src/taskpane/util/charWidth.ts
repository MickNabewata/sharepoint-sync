/** 半角変換 */
export function toHalfWidth(strVal) {
    if (strVal) {
        // 文字コードをシフト
        const halfVal = strVal.replace(
            /[！-～]/g,
            (tmpStr) => {
                return String.fromCharCode(tmpStr.charCodeAt(0) - 0xFEE0);
            }
        );

        // 文字コードシフトで対応できない文字の変換
        return halfVal
            .replace(/”/g, "\"")
            .replace(/’/g, "'")
            .replace(/‘/g, "`")
            .replace(/￥/g, "\\")
            .replace(/　/g, " ")
            .replace(/〜/g, "~");
    } else {
        return strVal;
    }
}