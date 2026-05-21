export class ErrorLogger {

  public static async logError(
    spContext: any,
    error: any,
    functionName: string,
    webPartName: string
  ): Promise<void> {

    try {
      const errorMessage = error?.message || JSON.stringify(error);
      const currentUser = await spContext.web.currentUser();

      await spContext.web.lists.getByTitle("ErrorLogs").items.add({
        Title: `${webPartName} - ${functionName}`,
        ErrorMessage: errorMessage,
        StackTrace: error?.stack || "",
        LoggedAt: new Date().toISOString(),
        PageUrl: window.location.href,
        LoggedBy: currentUser.Title 
      });

      console.log("Error logged successfully");

    } catch (e) {
      console.error("Logging failed:", e);
    }
  }
}