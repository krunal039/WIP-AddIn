import msalInstance from "./msalInstance";
import { loginRequest } from "./msalConfig";
import { AccountInfo, AuthenticationResult } from "@azure/msal-browser";

// Utility to get the current account
function getCurrentAccount(): AccountInfo | undefined {
  const accounts = msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : undefined;
}

// Get token for Microsoft Graph
export async function getGraphToken(): Promise<AuthenticationResult | null> {
  if (typeof msalInstance.initialize === "function") {
    await msalInstance.initialize();
  }
  const account = getCurrentAccount();
  const graphScopes = (process.env.REACT_APP_AZURE_GRAPH_SCOPES || "Mail.Send,Mail.ReadWrite,openid,profile,offline_access").split(",");
  
  // Debug account information
  console.log('MSAL Account:', account ? {
    username: account.username,
    localAccountId: account.localAccountId,
    homeAccountId: account.homeAccountId
  } : 'No account found');
  
  // Debug Office context
  try {
    const officeUserEmail = Office.context.mailbox.userProfile.emailAddress;
    console.log('Office context user:', officeUserEmail);
  } catch (error) {
    console.warn('Could not get Office context user:', error);
  }
  
  try {
    if (account) {
      return await msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account,
      });
    } else {
      throw new Error("No account available for silent token acquisition.");
    }
  } catch (silentError) {
    try {
      const result = await msalInstance.loginPopup({
        scopes: graphScopes,
      });
      msalInstance.setActiveAccount(result.account);
      return await msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account: result.account,
      });
    } catch (popupError) {
      return null;
    }
  }
}

// Get token for your API (.default)
export async function getApiToken(): Promise<AuthenticationResult | null> {
  if (typeof msalInstance.initialize === "function") {
    await msalInstance.initialize();
  }
  const account = getCurrentAccount();
  const apiScopes = (process.env.REACT_APP_AZURE_API_SCOPES || "api://d3398715-8435-43df-ac85-d28afd62f0e3/.default").split(",");
  try {
    if (account) {
      return await msalInstance.acquireTokenSilent({
        scopes: apiScopes,
        account,
      });
    } else {
      throw new Error("No account available for silent token acquisition.");
    }
  } catch (silentError) {
    try {
      const result = await msalInstance.loginPopup({
        scopes: apiScopes,
      });
      msalInstance.setActiveAccount(result.account);
      return await msalInstance.acquireTokenSilent({
        scopes: apiScopes,
        account: result.account,
      });
    } catch (popupError) {
      return null;
    }
  }
}