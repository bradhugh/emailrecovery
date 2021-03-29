import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import reportWebVitals from "./reportWebVitals";
import { store } from "./app/store";
import { Provider } from "react-redux";
import { Route, Switch } from "react-router";
import { Privacy } from "./Privacy";
import { Home } from "./Home";
import { BrowserRouter } from "react-router-dom";

const RootComponent: React.FC = () => {
  return (
    <React.StrictMode>
      <BrowserRouter>
        <Provider store={store}>
          <Switch>
            <Route path={[ "/privacy", "/apps/emailrecovery/privacy" ]}>
              <Privacy />
            </Route>
            <Route path={[ "/home", "/apps/emailrecovery/home" ]}>
              <Home />
            </Route>
            <Route path={[ "/", "/apps/emailrecovery" ]}>
              <App />
            </Route>
          </Switch>
        </Provider>
      </BrowserRouter>
    </React.StrictMode>
  );
};

ReactDOM.render(<RootComponent />, document.getElementById("root"));

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals();
