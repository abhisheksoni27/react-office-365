import React from "react";
import { UserAgentApplication } from "msal";

const defaultButtonText = "Sign in with Office 365";
const graphScopes = ["user.read"];

export default class OfficeLogin extends React.Component {
  constructor(props) {
    super(props);

    const applicationConfig = { clientID: props.clientID };

    const myMSALObj = new UserAgentApplication(
      applicationConfig.clientID,
      null,
      this.acquireTokenRedirectCallBack,
      {
        storeAuthStateInCookie: true,
        cacheLocation: "localStorage",
        redirectUri: "http://localhost:3000"
      }
    );

    this.state = { msal: myMSALObj };
  }

  acquireTokenRedirectCallBack = (errorDesc, token, error, tokenType) => {
    if (!token) {
      console.log(error + ":" + errorDesc);
    }
  };

  signIn = () => {
    const { msal } = this.state;

    msal
      .acquireTokenSilent(graphScopes)
      .then(accessToken => {})
      .catch(error => {
        return msal.acquireTokenPopup(graphScopes);
      })
      .then(accessToken => {})
      .catch(error => {
        return msal.loginPopup(graphScopes);
      })
      .then(accessToken => {});
    //AcquireTokenSilent Failure, send an interactive request.
  };

  handleClick = () => {
    this.signIn();
  };

  render() {
    const buttonText = this.props.text || defaultButtonText;
    return <button onClick={this.handleClick}>{buttonText}</button>;
  }
}
