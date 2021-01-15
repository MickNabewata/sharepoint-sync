import * as React from "react";
import ComponentBase from "../bases/ComponentBase";
import { IComboBoxOption, IComboBox, PrimaryButton } from "office-ui-fabric-react";
import Section from "../parts/Section";
import FullWidthComboBox from "../parts/FullWidthComboBox";

/** プロパティ型定義 */
export interface SourceSectionProps {
    /** Excelテーブル選択イベント */
    onChange: (value: Excel.Table) => void;
}

/** ステート型定義 */
export interface SourceSectionState {
    /** Excelテーブル選択肢一覧 */
    excelTables: IComboBoxOption[];
    /** 選択中のExcelテーブル */
    excelTableSelected: Excel.Table;
    /** コンポーネントが初期化済か否か */
    isComponentInitialized: boolean;
    /** Excelテーブル選択肢取得エラー */
    excelTablesErr: string;
    /** アニメーション実施 */
    animate: boolean;
}

/** インポート対象セクション コンポーネント */
export default class SourceSection extends ComponentBase<SourceSectionProps, SourceSectionState> {

    /** インポート対象セクション コンポーネント */
    constructor(props, context) {
        super(props, context);
        this.state = {
            excelTables: [],
            excelTableSelected: undefined,
            isComponentInitialized: false,
            excelTablesErr: "",
            animate: false
        };
    }

    /** Excelテーブルをコンボボックス選択肢形式に変換 */
    private toComboBoxOptions(tables: Excel.TableCollection): IComboBoxOption[] {
        const options: IComboBoxOption[] = [];

        // 選択肢を生成
        tables.items.forEach((table) => {
            // 選択肢を収集
            options.push({
                key: table.name,
                text: table.name
            });
        });

        return options.sort((v1, v2) => { return (v1.text < v2.text) ? -1 : 1; });
    }

    /** ファイル内のExcelテーブルをコンボボックス選択肢の形式ですべて返却 */
    private getTables(): Promise<IComboBoxOption[]> {
        try {
            // Excelファイル操作開始
            return Excel.run(async context => {
                // Excelテーブル一覧の読取
                const tables = context.workbook.tables.load();
                await context.sync().catch((ex) => { throw ex; });
                return this.toComboBoxOptions(tables);
            });
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** ファイル内のExcelテーブルをコンボボックス選択肢の形式でステートにセット */
    private async getTablesToState(): Promise<void> {
        try {
            await this.setToState({ isComponentInitialized: false });
            const tables = await this.getTables().catch((ex) => { throw ex; });
    
            const excelTablesErr = (tables && tables.length > 0) ? "" : "Excelテーブルがありません。";
            await this.setToState({ excelTables: tables, isComponentInitialized: true, excelTablesErr: excelTablesErr });
        } catch(ex) {
            await this.setToState({ isComponentInitialized: true, excelTablesErr: ex });
        }
    }

    /** アニメーション実行 */
    private animate(): Promise<void> {
        return new Promise<void>(async (resolve: () => void) => {
            await this.setToState({ animate: true });
            setTimeout(async () => {
                await this.setToState({ animate: false });
                resolve();
            }, 500);
        });
    }

    /** Excelテーブルの選択イベント */
    private handleExcelTableChanged = async (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        try {
            const { onChange } = this.props;

            console.log(event);
            console.log(option);
            console.log(index);
            console.log(value);
    
            // 処理中フラグ
            await this.setToState({ isComponentInitialized: false }); 
    
            // テーブル名をキーに１件特定
            await Excel.run(async context => {
                const tableRequest = context.workbook.tables.getItem(option.key.toString()).load();
                const table = await context.sync(tableRequest);
                table.getRange().select();
                await context.sync();
                await this.setToState({ excelTableSelected: table, isComponentInitialized: true });
            });
    
            onChange(this.state.excelTableSelected);
        } catch(ex) {
            await this.setToState({ excelTableSelected: undefined, isComponentInitialized: true, excelTablesErr: ex });
        }
    }

    /** 再読込ボタンクリックイベント */
    private handleRefreshButtonClicked = async () => {
        try {
            const { excelTableSelected } = this.state;
            const { onChange } = this.props;
    
            await this.getTablesToState().catch((ex) => { throw ex; });
            if (excelTableSelected) {
                const newOption = this.state.excelTables.filter((v) => { return v.key === excelTableSelected.name });
                if (newOption && newOption.length > 0) {
                    await Excel.run(async context => {
                        const table = context.workbook.tables.getItem(newOption[0].key.toString()).load();
                        await context.sync().catch((ex) => { throw ex; });
                        await this.setToState({ excelTableSelected: table }).catch((ex) => { throw ex; });
                    }).catch((ex) => { throw ex; });
                } else {
                    await this.setToState({ excelTableSelected: undefined });
                }
            }

            await this.animate().catch((ex) => { throw ex; });
            onChange(this.state.excelTableSelected);
        } catch(ex) {
            this.setToState({ excelTableSelected: undefined, excelTablesErr: ex });
        }
    }

    /** レンダリング */
    public render() {
        const { excelTables, excelTableSelected, isComponentInitialized, excelTablesErr, animate } = this.state;

        return (
            <Section title="インポート対象の選択">
                <FullWidthComboBox
                    placeholder="このファイル内のテーブルを選択します"
                    options={excelTables}
                    errorMessage={(excelTablesErr) ? excelTablesErr.toString() : ""}
                    disabled={!(excelTables && excelTables.length > 0) || !isComponentInitialized}
                    selectedKey={(excelTableSelected) ? excelTableSelected.name : undefined}
                    onChange={this.handleExcelTableChanged}
                    className={(animate === true) ? "ex-sp__animation-pulse" : undefined}
                />
                <PrimaryButton
                    text="再読込"
                    className="ex-sp__section-item"
                    onClick={this.handleRefreshButtonClicked}
                />
            </Section>
        );
    }

    /** コンポーネント描画完了後イベント */
    public componentDidMount() {
        // Excelテーブル名をすべて収集
        this.getTablesToState();
    }
}