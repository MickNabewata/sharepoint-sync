import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";

// アイコン初期化
initializeIcons();

let isOfficeInitialized = false;

/** レンダリング処理定義 */
const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Officeが初期化された後のイベントでrender */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* 初回render　この時点ではSpinnerだけ表示される */
render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
