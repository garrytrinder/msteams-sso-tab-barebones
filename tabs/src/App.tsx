
import { FluentProvider, Spinner } from '@fluentui/react-components';
import { HashRouter as Router, Redirect, Route } from "react-router-dom";
import { useTeamsFx } from "@microsoft/teamsfx-react";
import Tab from "./components/Tab";
import TabConfig from "./components/TabConfig";
import "./App.css";
import config from "./Config";
import { getThemeFromThemeString } from './helpers';
import { TeamsFxContext } from './context';

const App = () => {
  const { loading, themeString, teamsfx } = useTeamsFx({
    initiateLoginEndpoint: config?.initiateLoginEndpoint ?? '',
    clientId: config?.clientId ?? '',
  });

  return (
    <TeamsFxContext.Provider value={{ themeString, teamsfx }}>
      <FluentProvider theme={getThemeFromThemeString(themeString)}>
        <Router>
          <Route exact path="/">
            <Redirect to="/tab" />
          </Route>
          {loading ? (
            <Spinner style={{ margin: 100 }} />
          ) : (
            <>
              <Route exact path="/tab" component={Tab} />
              <Route exact path="/config" component={TabConfig} />
            </>
          )}
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}

export default App;
