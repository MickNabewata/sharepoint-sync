import * as React from "react";
import { IComboBoxOption, IComboBox, Text } from "office-ui-fabric-react";
import Section from "./Section";
import FullWidthComboBox from "./FullWidthComboBox";

/** マッピング */
export interface Map {
    /** Excel表フィールド名 */
    excelFieldName: string;
    /** SharePointリストフィールド名 */
    spoFieldName: string;
}

/** プロパティ型定義 */
export interface MapSectionProps {
    /** Excel表フィールド選択肢一覧 */
    excelFields: string[];
    /** SharePointリストフィールド選択肢一覧 */
    spoFields: IComboBoxOption[];
    /** 現在のマッピング一覧 */
    maps: Map[];
    /** マッピング変更イベント */
    onChange: (excelFieldName: string, spoFieldName: string) => void;
}

/** ステート型定義 */
export interface MapSectionState {
    /** コンポーネントが初期化済か否か */
    isComponentInitialized: boolean;
    /** エラーメッセージ */
    errorMessage: string;
}

/** マッピングセクション コンポーネント */
export default class MapSection extends React.Component<MapSectionProps, MapSectionState> {

    /** マッピングセクション コンポーネント */
    constructor(props, context) {
        super(props, context);
        this.state = {
            isComponentInitialized: false,
            errorMessage: ""
        };
    }

    /** ステートにセット */
    private setToState(state: Partial<MapSectionState>): Promise<void> {
        return new Promise<void>((resolve: () => void) => {
            this.setState(state as MapSectionState, () => {
                resolve();
            });
        });
    }

    /** SharePointリストフィールドの選択イベント */
    private handleSPOFieldChanged = (excelFieldName: string) => async (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        const { onChange } = this.props;

        console.log(event);
        console.log(option);
        console.log(index);
        console.log(value);

        // 変更イベント
        if (onChange) onChange(excelFieldName, (option && option.key) ? option.key.toString() : undefined);
    }

    /** 指定Excel表フィールドに対する現在のマッピングを取得 */
    private getCurrentMapping(excelFieldName: string) {
        const { maps } = this.props;
        const map = maps.find((v) => { return v.excelFieldName === excelFieldName; });
        return map ? map.spoFieldName : "";
    }

    /** レンダリング */
    public render() {
        const { isComponentInitialized, errorMessage } = this.state;
        const { excelFields, spoFields } = this.props;

        return (
            <Section title="フィールドマッピング">
                <div>
                    {
                        excelFields.map((excelField, i) => {
                            return (
                                <FullWidthComboBox
                                    key={`spo-field-${i}`}
                                    label={excelField}
                                    placeholder="このフィールドの取り込み先を選択します"
                                    options={spoFields}
                                    disabled={!(spoFields && spoFields.length > 0) || !isComponentInitialized}
                                    selectedKey={this.getCurrentMapping(excelField)}
                                    onChange={this.handleSPOFieldChanged(excelField)}
                                    className={i > 0 ? "ex-sp__section-item" : ""}
                                />
                            );
                        })
                    }
                </div>
                <div className="ex-sp__section-item">
                    <Text>{errorMessage}</Text>
                </div>
            </Section>
        );
    }

    /** コンポーネント描画完了後 */
    public componentDidMount() {
        this.setToState({ isComponentInitialized: true });
    }
}