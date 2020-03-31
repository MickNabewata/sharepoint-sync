import * as React from "react";
import ComponentBase from "./bases/ComponentBase";
import SourceSection from "./sections/SourceSection";
import Progress from "./parts/Progress";
import TargetSection from "./sections/TargetSection";
import MapSection, { Map } from "./sections/MapSection";
import ExecuteSection, { ExecuteStatus } from "./sections/ExecuteSection";
import Auth from "./auth/Auth";
import { Account } from "msal";
import { msalInstance } from "../pnp/mSal";
import { IComboBoxOption } from "office-ui-fabric-react";
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import { stringIsNullOrEmpty } from "@pnp/common";
import { initPnPJs } from "../pnp/pnp";

/** プロパティ型定義 */
export interface ExcelImporterProps {
    /** Officeが初期化済か否か */
    isOfficeInitialized: boolean;
}

/** ステート型定義 */
export interface ExcelImporterState {
    /** インポート対象 */
    selectedSource: Excel.Table;
    /** インポート対象のExcel表フィールド一覧 */
    selectedSourceFields: string[];
    /** インポート先のSharePointサイトURL */
    selectedWebUrl: string;
    /** インポート先のSharePointリストID */
    selectedTarget: string;
    /** インポート先のSharePointリストフィールド一覧 */
    selectedTargetFields: IComboBoxOption[];
    /** フィールドマッピング */
    maps: Map[];
    /** アプリケーションが初期化済か否か */
    isAppInitialized: boolean;
    /** 認証済みか否か */
    isAuthorized: boolean;
    /** 認証中のアカウント */
    account: Account;
    /** SharePointアクセストークン */
    token: string;
    /** SharePointテナントドメイン名 */
    domain: string;
    /** エラーメッセージ */
    err: string;
}

/** Excel Importer コンポーネント */
export default class SharePointSync extends ComponentBase<ExcelImporterProps, ExcelImporterState> {

    /** Excel Importer コンポーネント */
    constructor(props: ExcelImporterProps, context) {
        super(props, context);
        this.state = {
            selectedSource: undefined,
            selectedSourceFields: [],
            selectedWebUrl: undefined,
            selectedTarget: undefined,
            selectedTargetFields: [],
            maps: [],
            isAppInitialized: false,
            isAuthorized: false,
            account: undefined,
            token: undefined,
            domain: undefined,
            err: undefined
        };
    }

    /** インポート対象変更イベント */
    private handleSourceChanged = async (value: Excel.Table) => {
        const fields = await this.getExcelFields(value.name);
        await this.setToState({ selectedSource: value, selectedSourceFields: fields });
        await this.setToState({ maps: this.createMapping() });
    }

    /** インポート先変更イベント */
    private handleTargetChanged = async (webUrl: string, listId: string) => {
        const { token } = this.state;
        await this.setToState({ selectedWebUrl: webUrl, selectedTarget: listId });
        initPnPJs(sp, token, webUrl);
        const fields = await this.getSpoFields(listId);
        await this.setToState({ selectedTargetFields: fields });
        await this.setToState({ maps: this.createMapping() });
    }

    /** マッピング変更イベント */
    private handleMapChanged = (excelFieldName: string, spoFieldName: string) => {
        const newMaps = Array.from(this.state.maps);
        for (let i = 0; i < newMaps.length; i++) {
            if (newMaps[i].excelFieldName === excelFieldName) {
                newMaps[i].spoFieldName = spoFieldName;
                i = newMaps.length;
            }
        }
        this.setToState({ maps: newMaps });
    }

    /** 処理ステータス変更イベント */
    private handleExecuteStatusChanged = async (status: ExecuteStatus, message?: string): Promise<void> => {
        await this.setToState({ isAppInitialized: status === "completed", err: message });
    }

    /** ステートからマッピングを自動生成 */
    private createMapping(): Map[] {
        const { selectedSourceFields, selectedTargetFields } = this.state;

        let maps: Map[] = [];

        if (selectedSourceFields) {
            selectedSourceFields.forEach((sourceField) => {
                let targetField = selectedTargetFields.find((v) => { return (v.key.toString().toLocaleLowerCase() === sourceField.toLocaleLowerCase()); });
                if (!targetField) targetField = selectedTargetFields.find((v) => { return (v.text.toLocaleLowerCase().startsWith(sourceField.toLocaleLowerCase())); });
                maps.push({
                    excelFieldName: sourceField,
                    spoFieldName: (targetField && targetField.key) ? targetField.key.toString() : ""
                });
            });
        }

        return maps;
    }

    /** Excel表フィールド一覧を取得 */
    private async getExcelFields(tableName: string): Promise<string[]> {
        let fields: string[] = [];

        if (stringIsNullOrEmpty(tableName)) return fields;

        await Excel.run(async context => {
            const columns = context.workbook.tables.getItem(tableName).columns.load();
            await context.sync();
            columns.items.forEach((column) => {
                fields.push(column.name);
            });
        });

        fields = fields.sort((v1, v2) => { return (v1 < v2) ? -1 : 1; });

        return fields;
    }

    /** SharePointリストフィールド一覧を取得 */
    private async getSpoFields(listId: string): Promise<IComboBoxOption[]> {
        let options: IComboBoxOption[] = [];

        if (stringIsNullOrEmpty(listId)) return options;

        await sp.web.lists.getById(listId).fields.select("InternalName", "Title").filter("Sealed eq false").orderBy("Title", true).get().then(
            (fields) => {
                fields.forEach((field) => {
                    options.push({
                        key: field.InternalName,
                        text: `${field.Title} (${field.InternalName})`
                    });
                });
                return Promise.resolve();
            },
            (err) => {
                return this.setToState({ err: err });
            }
        );
        return options;
    }

    /** 認証成功イベント */
    private handleAuthorized = (account: Account, token: string, domain: string) => {
        this.setToState({ isAuthorized: true, account: account, token: token, domain: domain });
    }

    /** サインアウト */
    private handleSignOut = () => {
        this.setToState({ isAuthorized: false }).then(() => {
            msalInstance.logout();
        });
    }

    /** レンダリング */
    public render() {
        const { isOfficeInitialized } = this.props;
        const { selectedWebUrl, selectedSourceFields, selectedTargetFields, selectedSource, selectedTarget, maps, isAppInitialized, isAuthorized, account, token, domain, err } = this.state;

        return (
            <React.Fragment>
                {
                    (isAuthorized) ?
                        <React.Fragment>
                            <SourceSection onChange={this.handleSourceChanged} />
                            <TargetSection onChange={this.handleTargetChanged} account={account} token={token} domain={domain} singOutCallBack={this.handleSignOut} />
                            <MapSection excelFields={selectedSourceFields} spoFields={selectedTargetFields} maps={maps} onChange={this.handleMapChanged} />
                            <ExecuteSection token={token} webUrl={selectedWebUrl} listId={selectedTarget} tableName={selectedSource ? selectedSource.name : undefined} maps={maps} onStatusChanged={this.handleExecuteStatusChanged} />
                            <div>{err ? err.toString() : undefined}</div>
                            <Progress visible={(!isOfficeInitialized || !isAppInitialized)} />
                        </React.Fragment> :
                        <Auth authCallBack={this.handleAuthorized} logOutCallBack={this.handleSignOut} />
                }
            </React.Fragment>
        );
    }

    /** コンポーネント描画完了後イベント */
    public componentDidMount() {
        this.setToState({ isAppInitialized: true });
    }
}
