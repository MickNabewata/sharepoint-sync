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
            return Excel.run(context => {
                // Excelテーブル一覧の読取
                const tables = context.workbook.tables.load();
                return context.sync().then(
                    () => {
                        const options: IComboBoxOption[] = this.toComboBoxOptions(tables);

                        return Promise.resolve(options);
                    },
                    (err) => {
                        return Promise.reject(err);
                    }
                );
            });
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** ファイル内のExcelテーブルをコンボボックス選択肢の形式でステートにセット */
    private getTablesToState(): Promise<void> {
        return this.setToState({ isComponentInitialized: false }).then(
            () => {
                return this.getTables().then(
                    (tables) => {
                        // ステートにセット
                        const excelTablesErr = (tables && tables.length > 0) ? "" : "Excelテーブルがありません。";
                        return this.setToState({ excelTables: tables, isComponentInitialized: true, excelTablesErr: excelTablesErr });
                    },
                    (err) => {
                        // ステートにセット
                        return this.setToState({ isComponentInitialized: true, excelTablesErr: err });
                    }
                );
            },
            (err) => {
                // ステートにセット
                return this.setToState({ isComponentInitialized: true, excelTablesErr: err });
            }
        );
    }

    /** アニメーション実行 */
    private animate(): Promise<void> {
        return new Promise<void>((resolve: () => void) => {
            this.setToState({ animate: true }).then(() => {
                setTimeout(() => {
                    this.setToState({ animate: false }).then(
                        () => {
                            resolve();
                        }
                    );
                }, 500);
            });
        });
    }

    /** Excelテーブルの選択イベント */
    private handleExcelTableChanged = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        const { onChange } = this.props;

        console.log(event);
        console.log(option);
        console.log(index);
        console.log(value);

        // 処理中フラグ
        this.setToState({ isComponentInitialized: false }).then(
            () => {
                // テーブル名をキーに１件特定
                Excel.run(context => {
                    const table = context.workbook.tables.getItem(option.key.toString()).load();
                    return context.sync(table).then(
                        (table) => {
                            table.getRange().select();
                            return context.sync().then(
                                () => {
                                    return this.setToState({ excelTableSelected: table, isComponentInitialized: true }).then(
                                        () => {
                                            return Promise.resolve();
                                        }
                                    );
                                },
                                (err) => {
                                    return this.setToState({ excelTableSelected: undefined, isComponentInitialized: true, excelTablesErr: err });
                                }
                            );
                        },
                        (err) => {
                            return this.setToState({ excelTableSelected: undefined, isComponentInitialized: true, excelTablesErr: err });
                        }
                    );
                }).then(
                    () => {
                        onChange(this.state.excelTableSelected);
                    },
                    () => {
                        onChange(this.state.excelTableSelected);
                    }
                );
            }
        );
    }

    /** 再読込ボタンクリックイベント */
    private handleRefreshButtonClicked = () => {
        const { excelTableSelected } = this.state;
        const { onChange } = this.props;

        this.getTablesToState().then(
            () => {
                new Promise<void>((resolve: () => void, reject: (err) => void) => {
                    if (excelTableSelected) {
                        const newOption = this.state.excelTables.filter((v) => { return v.key === excelTableSelected.name });
                        if (newOption && newOption.length > 0) {
                            Excel.run(context => {
                                const table = context.workbook.tables.getItem(newOption[0].key.toString()).load();
                                return context.sync().then(
                                    () => {
                                        this.setToState({ excelTableSelected: table }).then(
                                            () => {
                                                resolve();
                                            },
                                            (err) => {
                                                reject(err);
                                            }
                                        );
                                    },
                                    (err) => {
                                        this.setToState({ excelTableSelected: undefined, excelTablesErr: err }).then(
                                            () => {
                                                reject(err);
                                            },
                                            () => {
                                                reject(err);
                                            }
                                        );
                                    }
                                );
                            });
                        } else {
                            this.setToState({ excelTableSelected: undefined }).then(
                                () => { resolve(); }
                            );
                        }
                    } else {
                        resolve();
                    }
                }).then(
                    () => {
                        this.animate();
                        onChange(this.state.excelTableSelected);
                    },
                    () => {
                        this.animate();
                        onChange(this.state.excelTableSelected);
                    }
                );
            }
        );
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