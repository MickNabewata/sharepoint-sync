import { toHalfWidth } from './charWidth';
import * as moment from 'moment';
import 'moment/locale/ja';
import { stringIsNullOrEmpty } from '@pnp/common';

/** 様々な型の値を文字列に変換 */
export function toString(value: string | number | Date): string {
    let ret = '';

    if (value !== null && value !== undefined) {
        switch (Object.prototype.toString.call(value)) {
            case '[object String]':
                ret = value.toString();
                break;
            case '[object Number]':
                ret = (value as Number).toLocaleString();
                break;
            case '[object Date]':
                ret = moment(value).format('YYYY/MM/DD HH:mm:ss');
                break;
            default:
                ret = value.toString();
                break;
        }
    }

    return ret;
}

/** 全角・半角が混在した文字列を半角数字に変換 変換失敗時はundefined */
export function toNumber(str: string): number {
    let ret = undefined;

    if (str) {
        // 半角変換して,を削除
        const halfValue: string = toHalfWidth(str).replace(/,/g, '');
        // 数値チェック
        const matches = halfValue.match(/^[+,-]?\d+(\.\d+)?$/g);
        if (matches && matches.length === 1 && matches[0].length === halfValue.length) {
            // 数値変換
            ret = Number(halfValue);
        }

    }

    return ret;
}

/** 全角・半角が混在した文字列が数値に変換可能かどうか判定 */
export function isNumber(str: string): boolean {
    let ret = false;

    if (str) {
        // 半角変換して,を削除
        const halfValue: string = toHalfWidth(str).replace(/,/g, '');
        // 数値チェック
        const matches = halfValue.match(/^[+,-]?\d+(\.\d+)?$/g);
        if (matches && matches.length === 1 && matches[0].length === halfValue.length) {
            ret = true;
        }

    }

    return ret;
}

/** 文字列を日付に変換 */
export function toDate(str: string): Date {
    if (stringIsNullOrEmpty(str)) return undefined;

    const m = moment(str);
    return m.isValid() ? m.toDate() : undefined;
}

/** 文字列を真偽値に変換 */
export function toBool(str: string): boolean {
    return str ? str.toLowerCase() === "true" : false;
}