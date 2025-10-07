import DebugService from "./DebugService";

const API_URL =
  process.env.REACT_APP_PLACEMENT_API_URL ||
  "https://localhost:4001/api/placements";
const SUBSCRIPTION_KEY = process.env.REACT_APP_PLACEMENT_API_KEY || "";

export interface PlacementRequestData {
  productCode: string;
  emailSender: string;
  emailSubject: string;
  emailReceivedDateTime: string;
  emlContent: string; // EML as string
}

export interface PlacementResponse {
  placementId: string;
  ingestionId: string;
  runId: string;
}

class PlacementApiService {
  private static instance: PlacementApiService;

  private constructor() {}

  public static getInstance(): PlacementApiService {
    if (!PlacementApiService.instance) {
      PlacementApiService.instance = new PlacementApiService();
    }
    return PlacementApiService.instance;
  }

  private getAuthHeaders(apiToken: string): Record<string, string> {
    return {
      "Ocp-Apim-Subscription-Key": SUBSCRIPTION_KEY,
      Authorization: `Bearer ${apiToken}`,
    };
  }

  public sanitizeEMLFileName(input: string): string {
  if (!input) return "email";

  // Lowercase
  let str = input.toLowerCase();

  // Replace invalid chars (anything not a-z, 0-9) with hyphen
  str = str.replace(/[^a-z0-9]/g, "-");

  // Replace multiple consecutive hyphens with single hyphen
  str = str.replace(/--+/g, "-");

  // Remove hyphen at start or end
  str = str.replace(/^-+/, "").replace(/-+$/, "");

  // Ensure starts with a letter (strip leading non-letters)
  str = str.replace(/^[^a-z]+/, "");

  // Ensure ends with a letter (strip trailing non-letters and digits)
  str = str.replace(/[^a-z]+$/, "");

  // If empty, default to "email"
  if (!str) {
    str = "email";
  }

  // Ensure length between 3 and 60
  if (str.length < 3) {
    str = str.padEnd(3, "a");
  } else if (str.length > 60) {
    str = str.substring(0, 60);
    // After trimming, re-strip trailing non-letters just in case
    str = str.replace(/[^a-z]+$/, "");
  }

  DebugService.debug("Original Subject:", input);
  DebugService.debug("Sanitize Subject:", str);

  return str;
}


  public async submitPlacementRequest(
    apiToken: string,
    data: PlacementRequestData
  ): Promise<PlacementResponse> {
    try {
      DebugService.placement("Starting placement request submission");
      DebugService.debug("EML content length:", data.emlContent.length);
      DebugService.debug(
        "EML content preview:",
        data.emlContent.substring(0, 300) + "..."
      );

      // Validate EML content before creating file
      if (!data.emlContent || data.emlContent.length === 0) {
        throw new Error("EML content is empty");
      }

      if (
        !data.emlContent.includes("From:") ||
        !data.emlContent.includes("Subject:")
      ) {
        throw new Error("EML content is missing required headers");
      }

      const emlBlob = new Blob([data.emlContent], { type: "message/rfc822" });
      const emlFile = new File(
        [emlBlob],
        `${this.sanitizeEMLFileName(data.emailSubject)|| "email"}.eml`,
        { type: "message/rfc822" }
      );

      DebugService.debug("EML file size:", emlFile.size);
      DebugService.debug("EML file name:", emlFile.name);
      DebugService.debug("EML file type:", emlFile.type);

      const formData = new FormData();
      formData.append("productCode", data.productCode);
      formData.append("emailSender", data.emailSender);
      formData.append("emailSubject", data.emailSubject);
      formData.append("emailReceivedDateTime", data.emailReceivedDateTime);
      formData.append("files", emlFile);

      const headers = this.getAuthHeaders(apiToken);

      DebugService.api("POST", API_URL, {
        productCode: data.productCode,
        emailSubject: data.emailSubject,
        emlFileSize: emlFile.size,
      });

      const response = await fetch(API_URL, {
        method: "POST",
        headers,
        body: formData,
      });

      if (!response.ok) {
        const error = await response.json();
        DebugService.error("Placement API failed", {
          status: response.status,
          error,
        });
        throw new Error(error.message || "Placement API call failed");
      }

      const result = await response.json();
      DebugService.placement(
        "Placement request submitted successfully",
        result
      );
      return result;
    } catch (error) {
      DebugService.errorWithStack(
        "Placement API request failed",
        error as Error
      );
      throw error;
    }
  }
}

export default PlacementApiService.getInstance();
