import type { ApiResponse, ScanLog, InventorySummaryItem, Product, User, ScanEvent, CakeStatus, LiveOperationsData, B2BClient } from '@shared/types';
// This is the production URL for the Google Apps Script backend.
const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzPS5aPW0x2fXmu2j1zcsSeo4eVi7fzERl63y0VX9qXEcWgKYZ9QGHf07jO2cBp-Uyf1Q/exec';
export class GoogleSheetClient {
  private apiKey: string;
  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }
  private async fetchFromGoogleScript<T>(
    method: 'GET' | 'POST',
    action: string,
    payload?: object
  ): Promise<ApiResponse<T>> {
    try {
      let response: Response;
      const headers: HeadersInit = {
        'Content-Type': 'application/json',
      };
      // Only add Authorization header if an API key is provided
      if (this.apiKey) {
        headers['Authorization'] = `Bearer ${this.apiKey}`;
      }
      if (method === 'GET') {
        const url = new URL(GOOGLE_SCRIPT_URL);
        url.searchParams.append('action', action);
        response = await fetch(url.toString(), {
          method: 'GET',
          headers,
        });
      } else { // POST
        response = await fetch(GOOGLE_SCRIPT_URL, {
          method: 'POST',
          headers,
          body: JSON.stringify({ action, payload }),
        });
      }
      const contentType = response.headers.get('content-type');
      if (response.ok && contentType && contentType.includes('application/json')) {
        return await response.json() as ApiResponse<T>;
      } else {
        const text = await response.text();
        const errorMessage = `Received non-JSON response from server: ${text}`;
        console.error(`Error calling Google Script action "${action}":`, errorMessage);
        return { success: false, error: errorMessage };
      }
    } catch (error) {
      console.error(`Error calling Google Script action "${action}":`, error);
      const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred while contacting the backend.';
      return { success: false, error: errorMessage };
    }
  }
  addScan = (serialNumber: string, scanEvent: ScanEvent, location: string, clientId?: string) => this.fetchFromGoogleScript<ScanLog>('POST', 'addScan', { serialNumber, scanEvent, location, clientId });
  getLogs = () => this.fetchFromGoogleScript<ScanLog[]>('GET', 'getLogs');
  clearLogs = () => this.fetchFromGoogleScript<{ message: string }>('POST', 'clearLogs');
  getSummary = () => this.fetchFromGoogleScript<InventorySummaryItem[]>('GET', 'getSummary');
  getProducts = () => this.fetchFromGoogleScript<Product[]>('GET', 'getProducts');
  addProduct = (product: Product) => this.fetchFromGoogleScript<Product[]>('POST', 'addProduct', product);
  deleteProduct = (productId: string) => this.fetchFromGoogleScript<Product[]>('POST', 'deleteProduct', { productId });
  getUsers = () => this.fetchFromGoogleScript<User[]>('GET', 'getUsers');
  getCakeStatus = () => this.fetchFromGoogleScript<CakeStatus[]>('GET', 'getCakeStatus');
  getLiveOperationsData = () => this.fetchFromGoogleScript<LiveOperationsData>('GET', 'getLiveOperationsData');
  getB2BClients = () => this.fetchFromGoogleScript<B2BClient[]>('GET', 'getB2BClients');
}
