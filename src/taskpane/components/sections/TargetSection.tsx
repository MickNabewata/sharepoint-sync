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
import { errToString } from "../../util/typeCheck";

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
    spoSitesErr: any;
    /** SharePointリスト選択肢取得エラー */
    spoListsErr: any;
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
    private async getWebs(): Promise<IComboBoxOption[]> {
        const { siteFilter } = this.state;
        const { token, domain } = this.props;

        try {
            // pnp初期化
            initPnPJs(sp, token, `https://${domain}.sharepoint.com`);

            // 検索条件
            const query: ISearchQueryBuilder =
                SearchQueryBuilder()
                    .text(
                        !stringIsNullOrEmpty(siteFilter) ?
                            `(contentclass:STS_Site OR contentclass:STS_Web) AND 
                                (
                                    Title:${siteFilter}* OR 
                                    Path:https://${domain}.sharepoint.com/sites/${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com/sites${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com/${siteFilter} OR
                                    Path:https://${domain}.sharepoint.com${siteFilter} OR
                                    Path:${siteFilter}
                                )` :
                            "(contentclass:STS_Site OR contentclass:STS_Web)")
                    .rowLimit(5000);
            
            // 検索
            const results = await sp.search(query).catch((ex) => { throw ex; });
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
            return options;
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointリストをコンボボックス選択肢の形式ですべて返却 */
    private async getLists(): Promise<IComboBoxOption[]> {
        const { token } = this.props;
        const { spoSiteSelected } = this.state;

        try {
            if (!spoSiteSelected) return [];

            // pnp初期化
            initPnPJs(sp, token, spoSiteSelected.key.toString());

            // 検索
            const lists = await sp.web.lists.select("Id", "Title", "IsSystemList").orderBy("Title", true).get().catch((ex) => { throw ex; });
            const options: IComboBoxOption[] = [];

            lists.forEach((list) => {
                if (list.IsSystemList === false) {
                    const listId = list.Id;
                    const listTitle = list.Title;
                    options.push({ key: listId, text: listTitle });
                }
            });

            return options;
        } catch (ex) {
            return Promise.reject(ex);
        }
    }

    /** SharePointサイトとリストをコンボボックス選択肢の形式でステートにセット */
    private async getWebsAndListsToState(): Promise<void> {
        let { spoSiteSelected, spoListSelected, spoLists } = this.state;
        let webs: IComboBoxOption[] = [];
        
        try {
            await this.setToState({ isComponentInitialized: false });
            webs = await this.getWebs().catch((ex) => { throw ex; });
            
            // 選択中のサイトをキーで再検索
            if(spoSiteSelected) {
                spoSiteSelected = webs.find((v) => { return v.key === spoSiteSelected.key; });

                // サイトが変わったらリストも取り直す
                spoLists = await this.getLists().catch((ex) => { throw ex; });

                // 選択中のリストをキーで再検索
                if (spoListSelected) {
                    spoListSelected = spoLists.find((v) => { return v.key === spoListSelected.key; });
                }
            }

            // ステートにセット
            const siteErr = (webs && webs.length > 0) ? "" : new Error("SharePointサイトがありません。");
            const listErr = (spoLists && spoLists.length > 0) ? "" : (spoSiteSelected) ? new Error("SharePointリストがありません。") : "";
            await this.setToState({ spoSites: webs, spoLists: spoLists, spoSiteSelected: spoSiteSelected, spoListSelected: spoListSelected, isComponentInitialized: true, spoSitesErr: siteErr, spoListsErr: listErr });
        } catch(ex) {
            await this.setToState({ spoSites: webs, spoLists: [], spoSiteSelected: spoListSelected, spoListSelected: undefined, isComponentInitialized: true, spoSitesErr: "", spoListsErr: ex });
        }
    }

    /** SharePointリストをコンボボックス選択肢の形式でステートにセット */
    private async getListsToState(): Promise<void> {
        let { spoSiteSelected, spoListSelected } = this.state;

        try {
            const lists = await this.getLists();

            // 選択中のリストをキーで再検索
            if (spoListSelected) {
                spoListSelected = lists.find((v) => { return v.key === spoListSelected.key; });
            }
            
            // ステートにセット
            const err = (spoSiteSelected && lists && lists.length > 0) ? "" : (spoSiteSelected) ? new Error("SharePointリストがありません。") : "";
            await this.setToState({ spoLists: lists, spoListSelected: spoListSelected, isComponentInitialized: true, spoListsErr: err });
        } catch(ex) {
            await this.setToState({ spoLists: [], spoListSelected: undefined, isComponentInitialized: true, spoListsErr: ex });
        }
    }

    /** アニメーション実行 */
    private animate(target: "site" | "list" | "both"): Promise<void> {
        return new Promise<void>(async (resolve: () => void) => {
            await this.setToState(this.createAnimateState(target, true));
            setTimeout(async () => {
                await this.setToState(this.createAnimateState(target, false));
                resolve();
              }, 500);
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
    private handleSiteChanged = async (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
        const { spoSites } = this.state;
        const { onChange } = this.props;

        console.log(event);
        console.log(option);
        console.log(index);
        console.log(value);

        // 選択肢から１件特定してステートにセット
        const newSiteUrl = option.key.toString();
        const selected = spoSites.find((v) => { return v.key === newSiteUrl; });
        await this.setToState({ spoSiteSelected: selected });

        // リストの選択肢を再取得
        await this.getListsToState();
        const listId = this.state.spoListSelected ? this.state.spoListSelected.key.toString() : undefined;
        await this.animate("list");
        onChange(newSiteUrl, listId);
    }

    /** SharePointリストの選択イベント */
    private handleListChanged = async (event: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
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
        await this.setToState({ spoListSelected: selected });
        onChange(webUrl, listId);
    }

    /** 再読込ボタンクリックイベント */
    private handleRefreshButtonClicked = async () => {
        
        const { onChange } = this.props;

        // SharePointサイトとリストを再取得
        await this.getWebsAndListsToState();
        
        // 変更を通知
        const { spoSiteSelected, spoListSelected } = this.state;
        const webUrl = spoSiteSelected ? spoSiteSelected.key.toString() : undefined;
        const listId = spoListSelected ? spoListSelected.key.toString() : undefined;
        this.animate("both");
        onChange(webUrl, listId);
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
                    errorMessage={(spoSitesErr) ? errToString(spoSitesErr) : ""}
                    disabled={!(spoSites && spoSites.length > 0) || !isComponentInitialized}
                    selectedKey={(spoSiteSelected) ? spoSiteSelected.key : undefined}
                    onChange={this.handleSiteChanged}
                    className={(animateSite === true) ? "ex-sp__section-item ex-sp__animation-pulse" : "ex-sp__section-item"}
                />
                <FullWidthComboBox
                    placeholder="SharePointリストを選択します"
                    options={spoLists}
                    errorMessage={(spoListsErr) ? errToString(spoListsErr) : ""}
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