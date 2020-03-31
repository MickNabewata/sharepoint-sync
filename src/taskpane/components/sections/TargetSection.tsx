import * as React from "react";
import ComponentBase from "../bases/ComponentBase";
import { IComboBoxOption, IComboBox, PrimaryButton, Text, TextField } from "office-ui-fabric-react";
import Section from "../parts/Section";
import FullWidthComboBox from "../parts/FullWidthComboBox";
import "@pnp/polyfill-ie11";
import { SearchQueryBuilder } from "@pnp/polyfill-ie11/dist/searchquerybuilder";
import { sp, ISearchQueryBuilder } from "@pnp/sp";
import { Account } from "msal";
import { stringIsNullOrEmpty } from "@pnp/common";
import { initPnPJs } from "../../pnp/pnp";

/** プロパティ型定義 */
export interface TargetSectionProps {
    /** SharePointリスト選択イベント */
    onChange: (webUrl: string, listId: string) => void;
    /** SharePointアクセストークン */
    token: string;
    /** 認証中のアカウント */
    account: Account;
    /** SharePointテナントドメイン名 */
    domain: string;
    /** サインアウト時のコールバック */
    singOutCallBack: () => void;
}

/** ステート型定義 */
export interface TargetSectionState {
    /** サイト選択肢フィルタ条件 */
    siteFilter: string;
    /** SharePointサイト選択肢一覧 */
    spoSites: IComboBoxOption[];
    /** 選択中のSharePointサイト */
    spoSiteSelected: IComboBoxOption;
    /** SharePointリスト選択肢一覧 */
    spoLists: IComboBoxOption[];
    /** 選択中のSharePointリスト */
    spoListSelected: IComboBoxOption;
    /** コンポーネントが初期化済か否か */
    isComponentInitialized: boolean;
    /** SharePointサイト選択肢取得エラー */
    spoSitesErr: string;
    /** SharePointリスト選択肢取得エラー */
    spoListsErr: string;
    /** サイトアニメーション実施 */
    animateSite: boolean;
    /** リストアニメーション実施 */
    animateList: boolean;
}

/** インポート先セクション コンポーネント */
export default class TargetSection extends ComponentBase<TargetSectionProps, TargetSectionState> {

    /** インポート先セクション コンポーネント */
    constructor(props: TargetSectionProps, context) {
        super(props, context);
        this.state = {
            siteFilter: "",
            spoSites: [],
            spoSiteSelected: undefined,
            spoLists: [],
            spoListSelected: undefined,
            isComponentInitialized: false,
            spoSitesErr: "",
            spoListsErr: "",
            animateSite: false,
            animateList: false
        };
    }

    /** SharePointサイトをコンボボックス選択肢の形式ですべて返却 */
    private getWebs(): Promise<IComboBoxOption[]> {
        const { siteFilter } = this.state;
        const { domain } = this.props;

        try {
            const query: ISearchQueryBuilder =
                SearchQueryBuilder()
                    .text(
                        !stringIsNullOrEmpty(siteFilter) ?
                            `contentclass:STS_Site AND 
                                (
                                    Title:${siteFilter}* OR 
                                    Path:https://${domain}.sharepoint.com/sites/${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com/sites${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com/${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com${siteFilter} OR
                                    Path:${siteFilter}
                                )` :
                            "contentclass:STS_Site")
                    .rowLimit(5000);
            return sp.search(query).then(
                async (results) => {
                    let options: IComboBoxOption[] = [];
                    let page = 1;
                    let gotoNext = true;
                    while (gotoNext) {
                        const result = await results.getPage(page, 100);
                        page++;
                        if (result && result.PrimarySearchResults && result.PrimarySearchResults.length > 0) {
                            result.PrimarySearchResults.forEach((result) => {
                                const siteName = result.Title;
                                const siteUrl = result.Path;
                                options.push({ key: siteUrl, text: `${siteName} (${siteUrl})` });
                            });
                        } else {
                            gotoNext = false;
                        }
                    }
                    options = options.sort((v1, v2) => { return (v1.text < v2.text) ? -1 : 1; });
                    return Promise.resolve(options);
                },
                (err) => {
                    return Promise.reject(err);
                }
            );
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointリストをコンボボックス選択肢の形式ですべて返却 */
    private getLists(): Promise<IComboBoxOption[]> {
        try {
            return sp.web.lists.select("Id", "Title", "IsSystemList").orderBy("Title", true).get().then(
                (lists) => {

                    const options: IComboBoxOption[] = [];

                    lists.forEach((list) => {
                        if (list.IsSystemList === false) {
                            const listId = list.Id;
                            const listTitle = list.Title;
                            options.push({ key: listId, text: listTitle });
                        }
                    });

                    return Promise.resolve(options);
                },
                (err) => {
                    return Promise.reject(err);
                }
            );
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointサイトとリストをコンボボックス選択肢の形式でステートにセット */
    private getWebsAndListsToState(): Promise<void> {
        let { spoSiteSelected, spoListSelected } = this.state;
        const { token } = this.props;

        return this.setToState({ isComponentInitialized: false }).then(
            () => {
                return this.getWebs().then(
                    (webs) => {
                        if (spoSiteSelected) {
                             // 選択中のサイトをキーで再検索
                            spoSiteSelected = webs.find((v) => { return v.key === spoSiteSelected.key; });

                            // PnP.jsを再初期化
                            initPnPJs(sp, token, spoSiteSelected.key.toString());

                            // サイトが変わったらリストも取り直す
                            return this.getLists().then(
                                (lists) => {
                                    // 選択中のリストをキーで再検索
                                    if (spoListSelected) {
                                        spoListSelected = lists.find((v) => { return v.key === spoListSelected.key; });
                                    }

                                    // ステートにセット
                                    const siteErr = (webs && webs.length > 0) ? "" : "SharePointサイトがありません。";
                                    const listErr = (lists && lists.length > 0) ? "" : (spoSiteSelected) ? "SharePointリストがありません。" : "";
                                    return this.setToState({ spoSites: webs, spoLists: lists, spoSiteSelected: spoSiteSelected, spoListSelected: spoListSelected, isComponentInitialized: true, spoSitesErr: siteErr, spoListsErr: listErr });
                                },
                                (err) => {
                                    // ステートにセット
                                    return this.setToState({ spoSites: webs, spoLists: [], spoSiteSelected: spoListSelected, spoListSelected: undefined, isComponentInitialized: true, spoSitesErr: "", spoListsErr: err });
                                }
                            );
                        } else {
                            // ステートにセット
                            return this.setToState({ spoSites: webs, spoLists: [], spoSiteSelected: undefined, spoListSelected: undefined, isComponentInitialized: true, spoSitesErr: "", spoListsErr: "" });
                        }
                    },
                    (err) => {
                        // ステートにセット
                        return this.setToState({ spoSites: [], spoLists: [], spoSiteSelected: undefined, spoListSelected: undefined, isComponentInitialized: true, spoSitesErr: err, spoListsErr: "" });
                    }
                );
            },
            (err) => {
                // ステートにセット
                return this.setToState({ spoSites: [], spoLists: [], spoSiteSelected: undefined, spoListSelected: undefined, isComponentInitialized: true, spoSitesErr: err, spoListsErr: "" });
            }
        );
    }

    /** SharePointリストをコンボボックス選択肢の形式でステートにセット */
    private getListsToState(): Promise<void> {
        let { spoSiteSelected, spoListSelected } = this.state;
        return this.setToState({ isComponentInitialized: false }).then(
            () => {
                return this.getLists().then(
                    (lists) => {
                        // 選択中のリストをキーで再検索
                        if (spoListSelected) {
                            spoListSelected = lists.find((v) => { return v.key === spoListSelected.key; });
                        }
                        
                        // ステートにセット
                        const err = (spoSiteSelected && lists && lists.length > 0) ? "" : (spoSiteSelected) ? "SharePointリストがありません。" : "";
                        return this.setToState({ spoLists: lists, spoListSelected: spoListSelected, isComponentInitialized: true, spoListsErr: err });
                    },
                    (err) => {
                        // ステートにセット
                        return this.setToState({ spoLists: [], spoListSelected: undefined, isComponentInitialized: true, spoListsErr: err });
                    }
                );
            },
            (err) => {
                // ステートにセット
                return this.setToState({ spoLists: [], spoListSelected: undefined, isComponentInitialized: true, spoListsErr: err });
            }
        );
    }

    /** アニメーション実行 */
    private animate(target: "site" | "list" | "both"): Promise<void> {
        return new Promise<void>((resolve: () => void) => {
            this.setToState(this.createAnimateState(target, true)).then(() => {
              setTimeout(() => {
                this.setToState(this.createAnimateState(target, false)).then(() => {
                  resolve();
                });
              }, 500);
            });
        });
    }

    /** アニメーション実施用ステートにセットする値を生成 */
    private createAnimateState(target: "site" | "list" | "both", animate: boolean): Partial<TargetSectionState> {
        let ret: Partial<TargetSectionState> = {};

        switch (target) {
            case "site":
                ret = { animateSite: animate };
                break;
            case "list":
                ret = { animateList: animate };
                break;
            case "both":
                ret = { animateSite: animate, animateList: animate };
                break;
        }

        return ret;
    }

    /** SharePointサイトの選択イベント */
    private handleSiteChanged = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        const { spoSites } = this.state;
        const { token, onChange } = this.props;

        console.log(event);
        console.log(option);
        console.log(index);
        console.log(value);

        // sp.jsを再度初期化
        const newSiteUrl = option.key.toString();
        initPnPJs(sp, token, newSiteUrl);

        // 選択肢から１件特定してステートにセット
        const selected = spoSites.find((v) => { return v.key === newSiteUrl; });
        this.setToState({ spoSiteSelected: selected }).then(
            () => {
                // リストの選択肢を再取得
                this.getListsToState().then(
                    () => {
                        const listId = this.state.spoListSelected ? this.state.spoListSelected.key.toString() : undefined;
                        this.animate("list");
                        onChange(newSiteUrl, listId);
                    },
                    () => {
                        const listId = this.state.spoListSelected ? this.state.spoListSelected.key.toString() : undefined;
                        this.animate("list");
                        onChange(newSiteUrl, listId);
                    }
                );
            }
        );
    }

    /** SharePointリストの選択イベント */
    private handleListChanged = (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        const { spoSiteSelected, spoLists } = this.state;
        const { onChange } = this.props;

        console.log(event);
        console.log(option);
        console.log(index);
        console.log(value);

        // 選択肢から１件特定してステートにセット
        const webUrl = spoSiteSelected ? spoSiteSelected.key.toString() : undefined;
        const listId = option ? option.key.toString() : undefined;
        const selected = spoLists.find((v) => { return v.key === listId; });
        this.setToState({ spoListSelected: selected }).then(
            () => {
                onChange(webUrl, listId);
            }
        );
    }

    /** 再読込ボタンクリックイベント */
    private handleRefreshButtonClicked = () => {
        
        const { onChange } = this.props;

        // SharePointサイトとリストを再取得
        this.getWebsAndListsToState().then(
            () => {
                // 変更を通知
                const { spoSiteSelected, spoListSelected } = this.state;
                const webUrl = spoSiteSelected ? spoSiteSelected.key.toString() : undefined;
                const listId = spoListSelected ? spoListSelected.key.toString() : undefined;
                this.animate("both");
                onChange(webUrl, listId);
            },
            () => {
                // 変更を通知
                const { spoSiteSelected, spoListSelected } = this.state;
                const webUrl = spoSiteSelected ? spoSiteSelected.key.toString() : undefined;
                const listId = spoListSelected ? spoListSelected.key.toString() : undefined;
                this.animate("both");
                onChange(webUrl, listId);
            }
        );
    }

    /** サインアウトボタンクリックイベント */
    private handleSignOutButtonClicked = () => {
        const { singOutCallBack } = this.props;
        if (singOutCallBack) singOutCallBack();
    }

    /** サイトの絞り込み条件入力イベント */
    private handleSiteFilterChanged = async (event: React.ChangeEvent<HTMLInputElement>, newValue: string) => {
        console.log(event);
        console.log(newValue);

        await this.setToState({ siteFilter: newValue });
        this.getWebsAndListsToState();
    }

    /** エラーオブジェクトからメッセージを取得 */
    private errToString(err: any): string {
        try {
            if (err && err.response && err.response._bodyText) {
                const internalError = JSON.parse(err.response._bodyText);
                return internalError ?
                    internalError["odata.error"] ?
                        internalError["odata.error"].message ?
                            internalError["odata.error"].message.value ?
                                internalError["odata.error"].message.value :
                                JSON.stringify(internalError["odata.error"].message) :
                            JSON.stringify(internalError["odata.error"]) :
                        JSON.stringify(internalError) :
                    JSON.stringify(err);
            }
            else {
                return err ?
                    err.description ?
                        err.description :
                        err.error_description ?
                            err.error_description :
                            JSON.stringify(err) :
                    undefined;
            }
        } catch (ex) {
            return err;
        }
    }

    /** レンダリング */
    public render() {
        const { account } = this.props;
        const { siteFilter, spoSites, spoLists, spoSiteSelected, spoListSelected, isComponentInitialized, spoSitesErr, spoListsErr, animateSite, animateList } = this.state;
        
        return (
            <Section title="インポート先の選択">
                <div>
                    <Text>アカウント：{(account) ? account.userName : "不明"}</Text>
                </div>
                <div className="ex-sp__section-item">
                    <TextField
                        placeholder="SharePointサイトの絞り込み条件"
                        value={siteFilter}
                        onChange={this.handleSiteFilterChanged}
                    />
                </div>
                <FullWidthComboBox
                    placeholder="SharePointサイトを選択します"
                    options={spoSites}
                    errorMessage={(spoSitesErr) ? this.errToString(spoSitesErr) : ""}
                    disabled={!(spoSites && spoSites.length > 0) || !isComponentInitialized}
                    selectedKey={(spoSiteSelected) ? spoSiteSelected.key : undefined}
                    onChange={this.handleSiteChanged}
                    className={(animateSite === true) ? "ex-sp__section-item ex-sp__animation-pulse" : "ex-sp__section-item"}
                />
                <FullWidthComboBox
                    placeholder="SharePointリストを選択します"
                    options={spoLists}
                    errorMessage={(spoListsErr) ? this.errToString(spoListsErr) : ""}
                    disabled={spoSiteSelected === undefined || !(spoLists && spoLists.length > 0) || !isComponentInitialized}
                    selectedKey={(spoListSelected) ? spoListSelected.key : undefined}
                    onChange={this.handleListChanged}
                    className={(animateList === true) ? "ex-sp__section-item ex-sp__animation-pulse" : "ex-sp__section-item"}
                />
                <div className="ex-sp__section-item">
                    <PrimaryButton
                        text="再読込"
                        onClick={this.handleRefreshButtonClicked}
                    />
                    <PrimaryButton
                        text="ログアウト"
                        className="ex-sp__section-column"
                        onClick={this.handleSignOutButtonClicked}
                    />
                </div>
            </Section>
        );
    }

    /** コンポーネント描画完了後イベント */
    public componentDidMount() {
        // SharePointサイトとリストをすべて収集
        this.getWebsAndListsToState();
    }
}