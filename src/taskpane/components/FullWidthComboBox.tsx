import * as React from "react";
import { ComboBox, IComboBoxProps } from "office-ui-fabric-react";

/** コンボボックス選択肢 */
export interface IFullWidthComboBoxOptions {
  /** キー */
  key: string;
  /** 表示文字列 */
  text: string;
}

/** コンボボックスプロパティ */
export interface FullWidthComboBoxProps extends IComboBoxProps {
}

/** コンボボックス */
export default function FullWidthComboBox(props: FullWidthComboBoxProps) {
  const properties = Object.assign({}, props);

  const options = properties.options.map((option) => {
    option.styles = {
      optionTextWrapper: {
        width: "83vw"
      }
    };
    return option;
  });

  properties.options = options;

  return (
    <ComboBox
      {...properties}
    />
  );
}