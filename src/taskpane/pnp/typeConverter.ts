import { toDate, toBool, toNumber, isNumber } from "../util/typeCheck";
import { stringIsNullOrEmpty } from "@pnp/common";
import { SPRest, SPBatch } from "@pnp/sp";

/**
 *  SharePointリストのフィールドの型
 *  (集計値, タスクの結果, 外部データには対応しない)
 */
export type SpoFieldType = "Text" | "Note" | "Choice" | "MultiChoice" | "Number" | "DateTime" | "Lookup" | "LookupMulti" | "Boolean" | "User" | "UserMulti" | "URL" | "TaxonomyFieldType" | "TaxonomyFieldTypeMulti";

/** SharePointに登録できる形式に変換 */
export async function toSpoType(sp: SPRest, batch: SPBatch, key: string, value: string, type: SpoFieldType, separator: string = ","): Promise<{ key: string, value: any }> {
    if (stringIsNullOrEmpty(key)) return undefined;

    switch (type) {
        case "MultiChoice":
            return { key: key, value: toMultiObject(!stringIsNullOrEmpty(value) ? value.split(separator) : []) };
        case "Number":
            if (!stringIsNullOrEmpty(value) && !isNumber(value)) {
                return Promise.reject(new Error(`${value}は有効な数値ではありません。`));
            } else {
                return { key: key, value: !stringIsNullOrEmpty(value) ? toNumber(value) : null };
            }
        case "Lookup":
            // TODO
            if (!stringIsNullOrEmpty(value) && !isNumber(value)) {
                return Promise.reject(new Error(`${value}は有効な数値ではありません。`));
            } else {
                return { key: `${key}Id`, value: !stringIsNullOrEmpty(value) ? toNumber(value) : null };
            }
        case "DateTime":
            if (stringIsNullOrEmpty(value)) return { key: key, value: null };
            const date = toDate(value);
            if (!date) throw new Error(`${value}は有効な日付ではありません。`);
            return { key: key, value: date.toISOString() };
        case "LookupMulti":
            // TODO
            return undefined;
        case "Boolean":
            if (stringIsNullOrEmpty(value)) return { key: key, value: null };
            switch (value.toLocaleLowerCase()) {
                case "はい":
                case "yes":
                    return { key: key, value: true };
                case "いいえ":
                case "no":
                    return { key: key, value: false };
                default:
                    return { key: key, value: toBool(value) };
            }
        case "User":
            if (stringIsNullOrEmpty(value)) return { key: `${key}Id`, value: null };
            const user = await retriveUser(sp, batch, value);
            return { key: `${key}Id`, value: key === "Editor" || key === "Author" ? user.LoginName : user.Id };
        case "UserMulti":
            if (stringIsNullOrEmpty(value)) return { key: `${key}Id`, value: toMultiObject([]) };
            const tasks = !stringIsNullOrEmpty(value) ? value.split(separator).map(async (v) => { return retriveUser(sp, batch, v); }) : [];
            const users = await Promise.all(tasks);
            return {
                key: `${key}Id`, value: toMultiObject(users.map((user) => {
                    return key === "Editor" || key === "Author" ? user.LoginName : user.Id;
            })) };
        case "URL":
            return { key: key, value: !stringIsNullOrEmpty(value) ? toUrlObject(value) : null };
        case "TaxonomyFieldType":
            /* TODO
            if (stringIsNullOrEmpty(value)) return { key: key, value: null };
            const tax = toTaxonomyObject(value, "");
            return { value: value, key: tax ? JSON.stringify(tax) : "" };*/
            return undefined;
        case "TaxonomyFieldTypeMulti":
            /* TODO
            if (stringIsNullOrEmpty(value)) return { key: key, value: null };
            const taxs = value.split(separator).map((v) => { return toTaxonomyObject(v, ""); });
            return { value: value, key: toString(taxs) };*/
            return undefined;
        default:
            return { key: key, value: !stringIsNullOrEmpty(value) ? value : null };
    }
}
/** ユーザーID取得 */
async function retriveUser(sp: SPRest, batch: SPBatch, value: string) {
    const user = await sp.web.inBatch(batch).ensureUser(value);
    return user ? user.data : null;
}

/** URLオブジェクト型に変換 */
function toUrlObject(value: string, separator: string = ",") {
    if (stringIsNullOrEmpty(value)) return undefined;

    const values = value.split(separator);
    if (values.length < 2) return undefined;

    return {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: values[1],
        Url: values[0]
    }
}

/** メタデータオブジェクト型に変換 */
/*
function toTaxonomyObject(termName: string, termGuid: string) {
    return {
        __metadata: { type: 'SP.Taxonomy.TaxonomyFieldValue' },
        Label: termName,
        TermGuid: termGuid,
        //WssId: -1
    };
}*/

/** 配列を文字列に変換 */
/*
function toString(array: any[]) {
    if (!array) return "[]";

    let results = Array.from(array);
    if (!results.find((v) => { return v !== null && v !== undefined })) results = undefined;

    return results ? JSON.stringify(results) : "[]";
}*/

/** 複数形に変換 */
function toMultiObject(value: any) {
    return { results: value };
}