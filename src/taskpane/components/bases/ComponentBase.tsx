import * as React from "react";

/** コンポーネント基底クラス */
export default class ComponentBase<IProps, IStates> extends React.Component<IProps, IStates> {

  /** ステートにセット(非同期) */
  protected async setToState(state: Partial<IStates>): Promise<void> {
    return new Promise<void>((resolve: () => void) => {
      this.setState(state as IStates, () => {
        resolve();
      });
    });
  }
}