import * as React from "react";
import ComponentBase from "./bases/ComponentBase";
import { HashRouter as Router, Switch, Route } from "react-router-dom";
import SharePointSync from "./SharePointSync";
import MsalRedirect from "./auth/MsalRedirect";
import Test from "./Test";

/** アプリケーションルート プロパティ */
export interface AppProps {
  /** Officeが初期化済か否か */
  isOfficeInitialized: boolean;
}

/** アプリケーションルート ステート */
export interface AppStates {}

/** アプリケーションルート  */
export default class App extends ComponentBase<AppProps, AppStates> {

  /** アプリケーションルート  */
  constructor(props, context) {
      super(props, context);
      this.state = {
      };
  }

  /** レンダリング */
  render() {
    const { isOfficeInitialized } = this.props;
    return (
      <Router>
        <Switch>
          <Route exact path="/login" component={() => { return <MsalRedirect />; }} />
          <Route exact path="/" component={() => { return <SharePointSync isOfficeInitialized={isOfficeInitialized} />; }} />
          <Route exact path="/test" component={() => { return <Test />; }} />
        </Switch>
      </Router>
    );
  }
}