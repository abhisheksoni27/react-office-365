import React from "react";
import { UserAgentApplication } from "msal";
import "./style.css";

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
        redirectUri: props.redirectURL
      }
    );

    this.state = { accessToken: null };
  }

  acquireTokenRedirectCallBack = (errorDesc, token, error, tokenType) => {
    if (!token) {
      throw err;
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
          this.props.onSuccess(accessToken);
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
        <button className="OfficeLoginButton" onClick={this.handleClick}>
          {buttonText}
        </button>
      </div>
    );
  }
}
