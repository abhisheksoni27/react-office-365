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
        redirectUri: "http://localhost:3000"
      }
    );

    this.state = { accessToken: null };
  }

  acquireTokenRedirectCallBack = (errorDesc, token, error, tokenType) => {
    if (!token) {
      console.log(
        "​OfficeLogin -> acquireTokenRedirectCallBack -> errorDesc",
        errorDesc
      );
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
          this.props.isLoading(false);
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
