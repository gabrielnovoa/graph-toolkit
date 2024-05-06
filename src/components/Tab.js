// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from "react";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { FileList, Person } from "@microsoft/mgt-react";
import { Button } from "@fluentui/react-components";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { CacheService } from "@microsoft/mgt";
import config from "./lib/config";
import "./App.css";
import "./Tab.css";

class Tab extends React.Component {
  constructor(props) {
    super(props);
    const cacheId = Providers.getCacheId();
    CacheService.clearCacheById(cacheId);

    this.state = {
      showLoginPage: undefined,
    };
  }

  async componentDidMount() {
    await this.initTeamsFx();
    await this.initGraphToolkit(this.credential, this.scope);
    await this.checkIsConsentNeeded();
  }

  async initGraphToolkit(credential, scope) {
    const provider = new TeamsFxProvider(credential, scope);
    Providers.globalProvider = provider;
  }

  async initTeamsFx() {
    this.credential = new TeamsUserCredential({
      initiateLoginEndpoint: config.initiateLoginEndpoint,
      clientId: config.clientId,
    });

    this.scope = ["User.Read", "User.ReadBasic.All", "Calendars.ReadWrite", "Files.ReadWrite.All", "Contacts.Read"];
  }

  async loginBtnClick() {
    try {
      await this.credential.login(this.scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
      this.setState({ showLoginPage: false });
    } catch (err) {
      if (err.message?.includes("CancelledByUser")) {
        const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
        err.message +=
          '\nIf you see "AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application" ' +
          "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
          `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
      }

      alert("Login failed: " + err);
      return;
    }
  }

  async checkIsConsentNeeded() {
    let consentNeeded = false;
    try {
      await this.credential.getToken(this.scope);
    } catch (error) {
      consentNeeded = true;
    }
    this.setState({
      showLoginPage: consentNeeded,
    });
    Providers.globalProvider.setState(consentNeeded ? ProviderState.SignedOut : ProviderState.SignedIn);
    return consentNeeded;
  }

  render() {
    return (
      <div>
        {this.state.showLoginPage === false && (
          <div>
            <div>
              <Person personQuery="me"></Person>
            </div>
            <div>
              <FileList></FileList>
            </div>
          </div>
        )}

        {this.state.showLoginPage === true && (
          <div>
            <h2>Please Authorize!</h2>
            <Button appearance="primary" onClick={() => this.loginBtnClick()}>
              Authorize
            </Button>
          </div>
        )}
      </div>
    );
  }
}
export default Tab;
