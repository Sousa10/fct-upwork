import msal
import requests
import os
from config import AZURE_SETTINGS

class OneDriveHelper:
    def __init__(self):
        self.token_cache = msal.SerializableTokenCache()
        self.app = msal.PublicClientApplication(
            AZURE_SETTINGS['client_id'],
            authority=AZURE_SETTINGS['authority'],
            token_cache=self.token_cache,
        )
        self.access_token = self._get_token()

    def _get_token(self):
        # Load token cache from file if it exists
        cache_file = 'token_cache.bin'
        if os.path.exists(cache_file):
            self.token_cache.deserialize(open(cache_file, 'r').read())

        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(AZURE_SETTINGS['scopes'], account=accounts[0])
            if 'access_token' in result:
                return result['access_token']

        # If no valid token is found, initiate device flow
        flow = self.app.initiate_device_flow(scopes=AZURE_SETTINGS['scopes'])
        if 'user_code' not in flow:
            raise Exception("Failed to create device flow. Error: {}".format(flow.get('error')))
        print(flow['message'])  # Instruct the user to visit the URL and enter the code
        result = self.app.acquire_token_by_device_flow(flow)
        if 'access_token' in result:
            # Save token cache to file
            open(cache_file, 'w').write(self.token_cache.serialize())
            return result['access_token']
        else:
            raise Exception("Could not acquire access token. Error: {}".format(result.get('error_description', 'No error description available')))

    def update_excel(self, file_path, worksheet_name, range_address, values):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
        
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

    def read_excel(self, file_path, worksheet_name, range_address):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
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
    
    def recalculate_workbook(self, file_path):
        """
        Triggers recalculation for the entire workbook.
        """
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/application/calculate"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        data = {"calculationType": "Full"}  # Options: Full, Recalculate
        response = requests.post(endpoint, headers=headers, json=data)
        # Handle 204 response
        if response.status_code == 204:
            return {"status": "Success", "message": "Recalculation completed with no response body."}
        
        # Handle other errors
        response.raise_for_status()
        return response.json()
    
    def list_root_files(self):
        endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        return response.json()
    
    def list_files(self, folder_path=""):
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}:/children"
        headers = {'Authorization': f'Bearer {self.access_token}'}
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        return response.json()
        # try:
        #     response = requests.get(endpoint, headers=headers)
        #     response.raise_for_status()
        #     items = response.json().get("value", [])
        #     files = [item for item in items if "file" in item]
        #     return files
        # except requests.exceptions.RequestException as e:
        #     print(f"Error: {e.response.text}")
        #     raise
    
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
