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
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  });

  await msalInstance.initialize();

  const accounts = msalInstance.getAllAccounts();
  const activeAccount = msalInstance.getActiveAccount() || accounts[0] || null;

  if (activeAccount) {
    msalInstance.setActiveAccount(activeAccount);
  }

  render(activeAccount);
}

elements.signInButton.addEventListener("click", async () => {
  clearError();
  elements.signInButton.disabled = true;

  try {
    const response = await msalInstance.loginPopup({
      scopes: ["openid", "profile", "email"],
      prompt: "select_account",
    });

    msalInstance.setActiveAccount(response.account);
    render(response.account);
  } catch (error) {
    if (error?.errorCode !== "user_cancelled") {
      console.error(error);
      showError("Microsoft sign-in did not complete. Check the app registration and redirect address.");
    }
  } finally {
    elements.signInButton.disabled = false;
  }
});

elements.signOutButton.addEventListener("click", async () => {
  clearError();
  const account = msalInstance.getActiveAccount();

  try {
    await msalInstance.logoutPopup({ account });
    msalInstance.setActiveAccount(null);
    render(null);
  } catch (error) {
    console.error(error);
    showError("Sign-out did not complete. Close the browser tab and try again.");
  }
});

initialize().catch((error) => {
  console.error(error);
  showError("The sign-in service could not be initialized.");
});
