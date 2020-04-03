import * as React from "react";
import ComponentBase from "../bases/ComponentBase";
import { PrimaryButton } from "office-ui-fabric-react";
import Section from "../parts/Section";
import { Map } from "./MapSection";
import "@pnp/polyfill-ie11";
import { sp, ListItemFormUpdateValue } from "@pnp/sp";
import { stringIsNullOrEmpty } from "@pnp/common";
import { SpoFieldType, toSpoType } from "../../pnp/typeConverter";
import { isNumber, toNumber, toDate as toDate2, toString } from "../../util/typeCheck";
import { toDate } from "../../excel/typeConverter";
import { initPnPJs } from "../../pnp/pnp";
import Err from "../parts/Error";

/** 実行ステータス */
export type ExecuteStatus = "processing" | "completed";

/** プロパティ型定義 */
export interface ExecuteSectionProps {
    /** SharePointアクセストークン */
    token: string;
    /** SharePointサイトURL */
    webUrl: string;
    /** SharePointリストID */
    listId: string;
    /** Excelテーブル名 */
    tableName: string;
    /** マッピング一覧 */
    maps: Map[];
    /** 実行ステータス変更イベント */
    onStatusChanged: (status: ExecuteStatus, message?: string) => Promise<void>;
}

/** Excelから読み取ったセルの値 */
type ExcelRawType = string | number | boolean;

/** SharePoint登録用のマッピングデータ */
interface MapData extends Map {
    /** Excelから読み取ったセルの値 */
    excelRawData?: ExcelRawType;
    /** SharePoint登録用の値 */
    spoField?: { key: string, value: any };
}

/** SharePointリストのフィールド定義 */
interface SpoField {
    /** 内部名 */
    InternalName: string;
    /** 表示名 */
    Title: string;
    /** 型 */
    TypeAsString: SpoFieldType;
};

/** ステート型定義 */
export interface ExecuteSectionState {
    /** エラーメッセージ */
    err: any;
}

/** 実行セクション コンポーネント */
export default class ExecuteSection extends ComponentBase<ExecuteSectionProps, ExecuteSectionState> {

    /** 実行セクション コンポーネント */
    constructor(props: ExecuteSectionProps, context) {
        super(props, context);
        this.state = {
            err: ""
        };
    }

    /** インポート実行が可能か否かを判定 */
    private canImport() {
        const { listId, tableName, maps } = this.props;

        // プロパティが揃っていること
        if (stringIsNullOrEmpty(listId) || stringIsNullOrEmpty(tableName) || maps === null || maps === undefined || maps.length <= 0) {
            return false;
        }

        // マッピングが最低１フィールド分行われていること
        const mapped = maps.find((map) => { return !stringIsNullOrEmpty(map.excelFieldName) && !stringIsNullOrEmpty(map.spoFieldName) });
        if (!mapped) {
            return false;
        }

        // 可能
        return true;
    }

    /** インポートを実行ボタンクリックイベント */
    private handleExcecuteButtonClicked = async () => {
        const { webUrl, token, onStatusChanged } = this.props;
        try {
            await this.setToState({ err: undefined });

            // 親コンポーネントに処理開始を通知
            if (onStatusChanged) await onStatusChanged("processing");

            // PnP初期化
            initPnPJs(sp, token, webUrl);

            // SharePointのフィールド定義を取得
            const spoFields = await this.getSpoListFieldTypes().catch((ex) => { throw ex; });

            // Excelテーブルのデータを取得
            const datas = await this.getExcelDatas(spoFields).catch((ex) => { throw ex; });

            // SharePointに登録
            await this.setToSpoList(datas).catch((ex) => { throw ex; });

            // 親コンポーネントに処理終了を通知
            if (onStatusChanged) await onStatusChanged("completed");
        } catch (ex) {
            await this.setToState({ err: ex });

            // 親コンポーネントに処理終了を通知
            if (onStatusChanged) await onStatusChanged("completed");
        }
    }

    /** SharePointのフィールド定義を取得 */
    private async getSpoListFieldTypes(): Promise<SpoField[]> {
        try {
            const { listId, maps } = this.props;
            let results: SpoField[] = [];

            // フィルタ文字列の作成
            let requestFields: string[] = [];
            maps.forEach((map) => { if (!stringIsNullOrEmpty(map.spoFieldName)) requestFields.push(`InternalName eq '${map.spoFieldName}'`); });
            const filter = requestFields.join(" or ");

            // SharePointへの問合せ
            await sp.web.lists.getById(listId).fields.select("InternalName", "Title", "TypeAsString").filter(filter).get().then(
                async (fields) => {
                    if (fields) {
                        fields.forEach((field: SpoField) => {
                            results.push({
                                InternalName: field.InternalName,
                                Title: field.Title,
                                TypeAsString: field.TypeAsString
                            });
                        });
                        return Promise.resolve();
                    }
                },
                async (err) => {
                    return Promise.reject(err);
                }
            );

            // 返却
            return results;
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** Excelテーブルのデータを取得 */
    private async getExcelDatas(spoFields: SpoField[]): Promise<MapData[][]> {
        let errs: any[] = [];
        let results: MapData[][] = [];

        try {
            const { tableName, maps } = this.props;
            let headers: string[] = [];
            let values: ExcelRawType[][] = [];

            /** Excelから値を読取 */
            await Excel.run(async context => {
                try {
                    const headRange = await context.workbook.tables.getItem(tableName).getHeaderRowRange().load("values");
                    const bodyRange = await context.workbook.tables.getItem(tableName).getDataBodyRange().load(["values", "valueTypes"]);
                    await context.sync();
                    if (headRange && headRange.values && bodyRange && bodyRange.values && headRange.values.length === 1 && bodyRange.values.length > 0) {
                        headers = headRange.values[0].map((v) => { return v ? v.toString() : v; });
                        values = bodyRange.values;
                    }
                    return Promise.resolve();
                } catch (ex) {
                    return Promise.reject(ex);
                }
            }).catch((ex) => { throw ex; });

            // 読み取った値を整形
            let batch = sp.web.createBatch();
            if (headers && headers.length > 0 && values && values.length > 0) {
                // Excel行に対する繰り返し
                values.forEach(async (row) => {
                    try {
                        // マッピングが出来ている列に対する繰り返し
                        let mapDatas: MapData[] = maps.filter((map) => { return (!stringIsNullOrEmpty(map.excelFieldName) && !stringIsNullOrEmpty(map.spoFieldName)); });
                        let newMap: MapData[] = [];
                        mapDatas.forEach(async (mapData) => {
                            try {
                                let result = Object.assign({}, mapData);

                                // Excelの生データを格納
                                const columnIndex = headers.findIndex((v) => { return v === result.excelFieldName; });
                                result.excelRawData = row.length > columnIndex ? row[columnIndex] : undefined;

                                const spoField = spoFields.find((v) => { return v.InternalName === result.spoFieldName });
                                if (spoField) {
                                    // SharePointの列が日付型かつExcelの生データが数値変換可能な値である場合、
                                    // Excelの生データをシリアル値と見做して変換
                                    let spoData = result.excelRawData ? result.excelRawData.toString() : undefined;
                                    if (spoField.TypeAsString === "DateTime" && isNumber(spoData)) {
                                        const dayData = toDate(toNumber(spoData));
                                        spoData = dayData ? dayData.toISOString() : undefined;
                                    }

                                    // SharePointに登録可能な型に変換
                                    result.spoField = await toSpoType(sp, batch, spoField.InternalName, spoData, spoField.TypeAsString).catch(async (ex) => {
                                        throw ex;
                                    });
                                }

                                newMap.push(result);
                                return Promise.resolve();
                            } catch (ex) {
                                errs.push(ex);
                                throw ex;
                            }
                        });

                        // 戻り値に追加
                        results.push(newMap);

                        return;
                    } catch (ex) {
                        errs.push(ex);
                        throw ex;
                    }
                });
            }

            // 整形に必要なSharePointデータを解決
            await batch.execute().catch((ex) => { throw ex; });
        } catch (ex) {
        }

        if (errs.length > 0) throw errs;

        return results;
    }

    /** SharePointリストにデータを登録 */
    private async setToSpoList(datas: MapData[][]): Promise<void> {
        try {
            const items = await this.setNoSystemFields(datas).catch((ex) => { throw ex; });
            return await this.updateSystemFields(items).catch((ex) => { throw ex; });
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointのシステム列以外への登録/更新 */
    private async setNoSystemFields(datas: MapData[][]): Promise<any[]> {
        try {
            if (!datas || datas.length === 0) return [];

            const { listId } = this.props;
            let items: any[] = [];

            // バッチ処理
            let batch = sp.web.createBatch();

            // データ登録
            let errs: any[] = [];
            datas.forEach(async (data) => {
                if (data) {
                    // 追加/更新用のリストアイテムデータを生成
                    let item = {} as any;
                    data.forEach((field) => {
                        if (field && field.spoField) {
                            item[field.spoField.key] = field.spoField.value as any;
                        }
                    });
                    let item2 = Object.assign({}, item);
                    item2.EditorId = undefined;
                    item2.AuthorId = undefined;
                    item2.Created = undefined;
                    item2.Modified = undefined;

                    // 追加 or 更新の判断
                    if (item.ID) {
                        // 更新
                        await sp.web.lists.getById(listId).items.getById(item2.ID).inBatch(batch).update(item2).catch((err) => { errs.push(err); });
                        items.push(item);
                    } else {
                        // 追加
                        const addResult = await sp.web.lists.getById(listId).items.inBatch(batch).add(item2).catch((err) => { errs.push(err); });
                        if (addResult) {
                            item.ID = addResult.data.ID;
                            items.push(item);
                        }
                    }
                }
            });
            await batch.execute().catch((ex) => { throw ex; });

            // 返却
            if (errs.length > 0) {
                return Promise.reject(errs);
            } else {
                return items;
            }
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointのシステム列を更新 */
    private async updateSystemFields(items: any[]): Promise<void> {
        try {
            if (!items || items.length === 0) return;

            const { listId } = this.props;

            // バッチ処理
            let batch = sp.web.createBatch();

            // データ更新
            let errs: any[] = [];
            items.forEach(async (item) => {
                try {
                    if (item) {
                        // 更新用データ作成
                        const values: ListItemFormUpdateValue[] = [];
                        if (item.EditorId) {
                            values.push({
                                FieldName: "Editor",
                                FieldValue: `[{ 'Key': '${item.EditorId}' }]`
                            });
                        }
                        if (item.AuthorId) {
                            values.push({
                                FieldName: "Author",
                                FieldValue: `[{ 'Key': '${item.AuthorId}' }]`
                            });
                        }
                        if (item.Modified) {
                            const d = toDate2(item.Modified);
                            values.push({
                                FieldName: "Modified",
                                FieldValue: d ? toString(d): item.Modified
                            });
                        }
                        if (item.Created) {
                            const d = toDate2(item.Created);
                            values.push({
                                FieldName: "Created",
                                FieldValue: d ? toString(d) : item.Created
                            });
                        }

                        // データ更新依頼
                        const results = await sp.web.lists.getById(listId).items.getById(item.ID).inBatch(batch).validateUpdateListItem(values).catch((err) => { throw new Error(err); });
                        if (results) {
                            results.forEach((result) => {
                                if (result.HasException === true) {
                                    errs.push(new Error(`${result.FieldName}:${result.ErrorMessage}`));
                                }
                            });
                        }
                    }
                } catch (ex) {
                    errs.push(ex);
                }
            });
            await batch.execute().catch((ex) => { throw new ex; });

            // 返却
            if (errs.length > 0) {
                return Promise.reject(errs);
            } else {
                return Promise.resolve();
            }
        } catch (ex) {
            return Promise.reject(ex);
        }
    }
    
    /** レンダリング */
    public render() {
        const { err } = this.state;

        return (
            <Section title="実行">
                <PrimaryButton
                    text="インポートを実行"
                    onClick={this.handleExcecuteButtonClicked}
                    disabled={!this.canImport()}
                />
                <Err err={err} />
            </Section>
        );
    }
}