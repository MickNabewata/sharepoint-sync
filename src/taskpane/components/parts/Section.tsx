import * as React from "react";
import { Separator } from "office-ui-fabric-react";

/** セクションプロパティ */
export interface SectionProps {
  /** セクションタイトル */
  title: string;
  /** 子コンポーネント */
  children?: JSX.Element | JSX.Element[] | string | never[];
}

/** セクション */
export default function Section(props: SectionProps) {
  return (
    <section>
      <h2 className="ex-sp__section-title">{props.title}</h2>
      <Separator />
      <div className="ex-sp__section-body">
        {props.children}
      </div>
    </section>
  );
}