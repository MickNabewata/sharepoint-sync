import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";

/** 待機中表示プロパティ */
export interface ProgressProps {
  /** 表示有無 */
  visible: boolean;
}

/** 待機中表示 */
export default function Progress(props: ProgressProps) {
  return (
    <React.Fragment>
      {
        (props.visible) ?
          <div className="ex-sp__progress-layer">
            <Spinner type={SpinnerType.large} className="ex-sp__progress" />
          </div> :
          undefined
      }
    </React.Fragment>
  );
}