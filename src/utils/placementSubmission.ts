// Utility for placement submission in Outlook add-in
// Calls workbenchService.submitPlacement and returns the result

export async function submitPlacement(
  apiToken: string,
  graphToken: string,
  item: any,
  selectedProduct: string,
  sendCopyToCyberAdmin: boolean,
  workbenchService: any
): Promise<any> {
  return workbenchService.submitPlacement(
    apiToken,
    graphToken,
    item,
    selectedProduct,
    sendCopyToCyberAdmin
  );
} 