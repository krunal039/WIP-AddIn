import DebugService from "./DebugService";
import { environment } from "../config/environment";

export interface PlacementRequestData {
    productCode: string;
    searchString: string;
}

export interface PlacementResponse {
    placementId: string;
    insuredName: string;
    broker: string;
    brokerCode: string;
}

class PlacementApiSearchService {
    private static instance: PlacementApiSearchService;

    private constructor() { }

    public static getInstance(): PlacementApiSearchService {
        if (!PlacementApiSearchService.instance) {
            PlacementApiSearchService.instance = new PlacementApiSearchService();
        }
        return PlacementApiSearchService.instance;
    }

    /**
     * Gets the API URL dynamically from environment config
     */
    private getApiUrl(): string {
        //const url = "https://localhost:3005/api/placements";
        const url = environment.PLACEMENT_API_URL || "https://localhost:4001/api/placements";
        DebugService.debug("PlacementApiSearchService API_URL:", url);
        return url;
    }

    /**
     * Gets the subscription key dynamically from environment config
     */
    private getSubscriptionKey(): string {
        const key = environment.PLACEMENT_API_KEY || "";
        if (!key) {
            DebugService.warn("PlacementApiSearchService: PLACEMENT_API_KEY is empty");
        }
        return key;
    }

    private getAuthHeaders(apiToken: string): Record<string, string> {
        return {
            "Ocp-Apim-Subscription-Key": this.getSubscriptionKey(),
            Authorization: `Bearer ${apiToken}`,
        };
    }

    public async searchPlacementID(
        apiToken: string,
        data: PlacementRequestData
    ): Promise<PlacementResponse[]> {
        try {
            debugger;
            DebugService.placement("Starting placement search request");

            const queryParams = new URLSearchParams({
            productCode: data.productCode,
            searchString: data.searchString,
        }).toString();

            const headers = this.getAuthHeaders(apiToken);
            //const apiUrl = this.getApiUrl();
            const apiUrl = `${this.getApiUrl()}?${queryParams}`;

            DebugService.api("GET", apiUrl, {
                productCode: data.productCode,
                searchString: data.searchString,
            });
            console.log(headers);
            console.log(data.productCode);
            console.log(data.searchString);
            console.log(apiUrl);

            const response = await fetch(apiUrl, {
                method: "GET",
                headers,
            });

            if (!response.ok) {
                const error = await response.json();
                DebugService.error("Placement Search API failed", {
                    status: response.status,
                    error,
                });
                console.log(error);
                throw new Error(error.message || "Placement Search API call failed");
            }

            const result: PlacementResponse[] = await response.json();

            DebugService.placement(
                "Placement search request successful",
                result
            );
            return result;
        } catch (error) {
            DebugService.errorWithStack(
                "Placement Search API request failed",
                error as Error
            );
            throw error;
        }
    }
}

export default PlacementApiSearchService.getInstance();
