import * as React from "react";
import { errToString } from "../../util/typeCheck";

/** エラーメッセージ表示プロパティ */
export interface ErrorProps {
  /** エラーオブジェクト */
  err: any;
}

/** エラーメッセージ表示 */
export default function Error(props: ErrorProps) {
  const { err } = props;
  return (
    <React.Fragment>
      {err ? <div className="ex-sp__section-item">{errToString(err)}</div> : undefined}
    </React.Fragment>
  );
}