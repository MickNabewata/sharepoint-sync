import * as React from "react";
import { toDate } from "../excel/typeConverter";

/** テスト プロパティ */
export interface TestProps {
}

/** テスト ステート */
export interface TestStates {}

/** テスト  */
export default class App extends React.Component<TestProps, TestStates> {

  /** テスト  */
  constructor(props, context) {
      super(props, context);
      this.state = {
      };
  }

  render() {
    return (
      <div>
        <div>TEST</div>
        <div>2019/12/05:{toDate(43804).toISOString()}</div>
        <div>2021/01/05:{toDate(44201).toISOString()}</div>
      </div>
    );
  }
}