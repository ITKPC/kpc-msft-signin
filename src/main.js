import { PublicClientApplication } from "@azure/msal-browser";
import { entraConfig } from "./config.js";
import "./styles.css";

const elements = {
  setupWarning: document.querySelector("#setup-warning"),
  signedOutView: document.querySelector("#signed-out-view"),
  signedInView: document.querySelector("#signed-in-view"),
  signInButton: document.querySelector("#sign-in-button"),
  signOutButton: document.querySelector("#sign-out-button"),
  displayName: document.querySelector("#display-name"),
  username: document.querySelector("#username"),
  errorMessage: document.querySelector("#error-message"),
};

const isConfigured =
  entraConfig.clientId &&
  entraConfig.tenantId &&
  !entraConfig.clientId.startsWith("PASTE_") &&
  !entraConfig.tenantId.startsWith("PASTE_");

let msalInstance;

function showError(message) {
  elements.errorMessage.textContent = message;
  elements.errorMessage.hidden = false;
}

function clearError() {
  elements.errorMessage.textContent = "";
  elements.errorMessage.hidden = true;
}

function render(account) {
  const signedIn = Boolean(account);
  elements.signedOutView.hidden = signedIn;
  elements.signedInView.hidden = !signedIn;

  if (account) {
    elements.displayName.textContent = account.name || "Microsoft 365 user";
    elements.username.textContent = account.username || "Unknown account";
  }
}

async function initialize() {
  if (!isConfigured) {
    elements.setupWarning.hidden = false;
    elements.signInButton.disabled = true;
    return;
  }

  msalInstance = new PublicClientApplication({
    auth: {
      clientId: entraConfig.clientId,
      authority: `https://login.microsoftonline.com/${entraConfig.tenantId}`,
      redirectUri: window.location.origin,
      postLogoutRedirectUri: window.location.origin,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  });

  await msalInstance.initialize();

  const redirectResponse = await msalInstance.handleRedirectPromise();
  const account =
    redirectResponse?.account ||
    msalInstance.getActiveAccount() ||
    msalInstance.getAllAccounts()[0] ||
    null;

  if (account) {
    msalInstance.setActiveAccount(account);
  }

  render(account);
}

elements.signInButton.addEventListener("click", async () => {
  clearError();
  elements.signInButton.disabled = true;

  try {
    await msalInstance.loginRedirect({
      scopes: ["openid", "profile", "email"],
      prompt: "select_account",
    });
  } catch (error) {
    console.error(error);
    const detail = error?.errorCode || error?.message || "unknown_error";
    showError(`Microsoft sign-in did not complete: ${detail}`);
    elements.signInButton.disabled = false;
  }
});

elements.signOutButton.addEventListener("click", async () => {
  clearError();
  const account = msalInstance.getActiveAccount();

  try {
    await msalInstance.logoutRedirect({ account });
  } catch (error) {
    console.error(error);
    const detail = error?.errorCode || error?.message || "unknown_error";
    showError(`Sign-out did not complete: ${detail}`);
  }
});

initialize().catch((error) => {
  console.error(error);
  const detail = error?.errorCode || error?.message || "unknown_error";
  showError(`The sign-in service could not be initialized: ${detail}`);
});
