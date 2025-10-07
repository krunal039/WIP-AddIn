// Outlook notification banner utility

export async function showWorkbenchNotificationBanner(
  item: any, 
  workbenchId: string, 
  DebugService: any
): Promise<void> {
  if (!item || !workbenchId) {
    DebugService.warn('Cannot show notification banner: missing item or workbenchId');
    return;
  }

  DebugService.debug('Showing Outlook notification banner with WBID:', workbenchId);

  return new Promise((resolve) => {
    // Check if notificationMessages API is available
    if (!item.notificationMessages || typeof item.notificationMessages.replaceAsync !== 'function') {
      DebugService.warn('notificationMessages API not available on this item');
      resolve();
      return;
    }

    // Use the exact format provided by the user
    item.notificationMessages.replaceAsync(
      "workbenchSubmissionBanner", // Unique key for this message
      {
        type: "informationalMessage",
        message: `This email has been successfully submitted for ingestion with Placement ID: ${workbenchId}`,
        icon: "Icon.16x16", // Icon from manifest assets
        persistent: true // If true, stays until dismissed
      },
      function (asyncResult: Office.AsyncResult<void>) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          DebugService.debug('Banner displayed successfully!');
        } else {
          DebugService.error('Failed to display banner:', asyncResult.error?.message);
        }
        resolve();
      }
    );
  });
}