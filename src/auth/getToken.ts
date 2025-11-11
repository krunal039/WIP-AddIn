import { getMsalInstance } from "./msalInstance";
import { getLoginRequest } from "./msalConfig";
import { AccountInfo, AuthenticationResult } from "@azure/msal-browser";
import { environment } from "../config/environment";
import DebugService from "../service/DebugService";

// Utility to get the current account
function getCurrentAccount(): AccountInfo | undefined {
  const msalInstance = getMsalInstance();
  const accounts = msalInstance.getAllAccounts();
  return accounts.length > 0 ? accounts[0] : undefined;
}

// Get token for Microsoft Graph
export async function getGraphToken(): Promise<AuthenticationResult | null> {
  const msalInstance = getMsalInstance();
  if (typeof msalInstance.initialize === "function") {
    await msalInstance.initialize();
  }
  const account = getCurrentAccount();
  const graphScopes = environment.AZURE_GRAPH_SCOPES.split(",").map(s => s.trim());
  
  // Debug account information
  DebugService.debug('MSAL Account:', account ? {
    username: account.username,
    localAccountId: account.localAccountId,
    homeAccountId: account.homeAccountId
  } : 'No account found');
  
  // Debug Office context
  try {
    const officeUserEmail = Office.context.mailbox.userProfile.emailAddress;
    DebugService.debug('Office context user:', officeUserEmail);
  } catch (error) {
    DebugService.warn('Could not get Office context user:', error);
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
  const msalInstance = getMsalInstance();
  if (typeof msalInstance.initialize === "function") {
    await msalInstance.initialize();
  }
  const account = getCurrentAccount();
  const apiScopes = environment.AZURE_API_SCOPES.split(",").map(s => s.trim());
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