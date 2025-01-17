import msal
import requests
from config import AZURE_SETTINGS

class OneDriveHelper:
    def __init__(self):
        self.app = msal.PublicClientApplication(
            AZURE_SETTINGS['client_id'],
            authority=AZURE_SETTINGS['authority'],
            token_cache=msal.SerializableTokenCache(),
            # client_credential=AZURE_SETTINGS['client_secret']
        )
        self.access_token = self._get_token()

    def _get_token(self):
        result = self.app.acquire_token_interactive(AZURE_SETTINGS['scopes'])
        if not result:
            result = self.app.acquire_token_for_client(scopes=AZURE_SETTINGS['scopes'])
        if 'access_token' in result:
            return result['access_token']
        else:
            raise Exception("Could not acquire access token. Error: {}".format(result.get('error_description', 'No error description available')))

    def update_excel(self, file_id, worksheet_name, range_address, values):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        data = {
            'values': values
        }
        # print(f"Endpoint: {endpoint}")
        # print(f"Payload: {data}")
        response = requests.patch(endpoint, headers=headers, json=data)
        # print(f"Response: {response.status_code}, {response.text}")
        response.raise_for_status()
        return response.json()

    def read_excel(self, file_id, worksheet_name, range_address):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        return response.json()
    
    def list_drives(self):
        endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        try:
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error: {e.response.text}")
            raise
    
    def recalculate_workbook(self, file_id):
        """
        Triggers recalculation for the entire workbook.
        """
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/application/calculate"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        data = {"calculationType": "Full"}  # Options: Full, Recalculate
        response = requests.post(endpoint, headers=headers, json=data)
        # Handle 204 response
        if response.status_code == 204:
            return {"status": "Success", "message": "Recalculation completed with no response body."}
        
        # Handle other errors
        response.raise_for_status()
        return response.json()
    
    def list_files(self):
        endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        try:
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            items = response.json().get("value", [])
            files = [item for item in items if "file" in item]
            return files
        except requests.exceptions.RequestException as e:
            print(f"Error: {e.response.text}")
            raise
    
    def get_file_metadata(self, file_path):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/"
        headers = {'Authorization': f'Bearer {self.access_token}'}

        try:
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()  # Raise an exception for HTTP errors
            return response.json()  # Return the file metadata as a dictionary
        except requests.exceptions.RequestException as e:
            print(f"Error: {e.response.text if e.response else str(e)}")
            raise

