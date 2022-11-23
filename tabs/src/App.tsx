import { FluentProvider } from "@fluentui/react-components";
import { HashRouter as Router, Redirect, Route } from "react-router-dom";
import { useTeamsFx } from "@microsoft/teamsfx-react";
import Tab from "./components/Tab";
import TabConfig from "./components/TabConfig";
import "./App.css";
import config from "./Config";
import { getThemeFromThemeString } from "./helpers";
import { TeamsFxContext } from "./context";
import { app } from "@microsoft/teams-js";
import { useEffect } from "react";

const App = () => {
  const { loading, themeString, teamsfx } = useTeamsFx({
    initiateLoginEndpoint: config?.initiateLoginEndpoint ?? "",
    clientId: config?.clientId ?? "",
  });

  useEffect(() => {
    app.initialize().then(() => {
      app.notifySuccess();
    });
  });

  useEffect(() => {
    if (!loading) {
      app.notifyAppLoaded();
    }
  }, [loading]);

  return (
    <TeamsFxContext.Provider value={{ themeString, teamsfx }}>
      <FluentProvider theme={getThemeFromThemeString(themeString)}>
        <Router>
          <Route exact path="/">
            <Redirect to="/tab" />
          </Route>
          <Route exact path="/tab" component={Tab} />
          <Route exact path="/config" component={TabConfig} />
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
};

export default App;
