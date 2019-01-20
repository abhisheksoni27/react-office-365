import React from "react";
import { UserAgentApplication } from "msal";
import MicrosoftLogo from "./microsoft.svg";
const defaultButtonText = "Sign in with Office 365";
const graphScopes = ["user.read"];

export default class OfficeLogin extends React.Component {
  constructor(props) {
    super(props);

    const applicationConfig = { clientID: props.clientID };

    this.msal = new UserAgentApplication(
      applicationConfig.clientID,
      null,
      this.acquireTokenRedirectCallBack,
      {
        storeAuthStateInCookie: true,
        cacheLocation: "localStorage",
        redirectUri: "http://localhost:3000"
      }
    );

    this.state = { accessToken: null };
  }

  acquireTokenRedirectCallBack = (errorDesc, token, error, tokenType) => {
    if (!token) {
      console.log(error + ":" + errorDesc);
    }
  };

  signIn = () => {
    this.msal
      .loginPopup(graphScopes)
      .then(idToken => {
        return this.msal.acquireTokenSilent(graphScopes);
      })
      .then(accessToken => {
        if (accessToken) {
          this.props.isLoading(false);
          this.props.onSuccess(accessToken, msal.getUser().name);
        }
      })
      .catch(err => {
        this.props.onFailure(err);
      });
  };

  handleClick = () => {
    this.props.isLoading(true);
    this.signIn();
  };

  render() {
    const buttonText = this.props.text || defaultButtonText;
    return (
      <div>
        <button onClick={this.handleClick}>{buttonText}</button>
      </div>
    );
  }
}
