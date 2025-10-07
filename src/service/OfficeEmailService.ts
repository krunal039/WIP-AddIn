import { EmailMetadata } from '../types';

export class OfficeEmailService {
  private static instance: OfficeEmailService;

  private constructor() {}

  public static getInstance(): OfficeEmailService {
    if (!OfficeEmailService.instance) {
      OfficeEmailService.instance = new OfficeEmailService();
    }
    return OfficeEmailService.instance;
  }

  public async showNotification(message: string, type: 'success' | 'error' | 'info' = 'info'): Promise<void> {
    return new Promise((resolve) => {
      if (!Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        resolve();
        return;
      }
      try {
        const item = Office.context.mailbox.item;
        //item.notificationMessages.addAsync(
        //  `notification-${Date.now()}`,
        //  {
        //    type: type === 'error' ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage :
        //           type === 'success' ? Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage :
        //           Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        //    message: message
        //  },
        //  () => setTimeout(resolve, 5000)
        //);
      } catch (error) {
        resolve();
      }
    });
  }

  public isSupported(): boolean {
    return !!(Office && Office.context && Office.context.mailbox && Office.context.mailbox.item);
  }

  public getOfficeContext(): any {
    return Office.context;
  }
} 