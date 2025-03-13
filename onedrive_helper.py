import msal
import requests
import os
from config import AZURE_SETTINGS
import re

class OneDriveHelper:
    def __init__(self):
        self.token_cache = msal.SerializableTokenCache()
        self.app = msal.PublicClientApplication(
            AZURE_SETTINGS['client_id'],
            authority=AZURE_SETTINGS['authority'],
            token_cache=self.token_cache,
        )
        self.session = requests.Session()  #FIXED: Initialize session
        self.access_token = self._get_token()
        # self.check_token_permissions(self.access_token)

    def refresh_token(self):
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(AZURE_SETTINGS['scopes'], account=accounts[0])
            if 'access_token' in result:
                open('token_cache.bin', 'w').write(self.token_cache.serialize())
                return result['access_token']
        return self._get_token()

    def _get_token(self):
        # Load token cache from file if it exists
        cache_file = 'token_cache.bin'
        if os.path.exists(cache_file):
            self.token_cache.deserialize(open(cache_file, 'r').read())

        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(AZURE_SETTINGS['scopes'], account=accounts[0])
            if result and 'access_token' in result:
                print("🔄 Using cached access token")
                return result['access_token']

        # If no valid token is found, initiate device flow
        print("⚡ Requesting new access token via Device Code Flow...")
        flow = self.app.initiate_device_flow(scopes=AZURE_SETTINGS['scopes'])
        if 'user_code' not in flow:
            raise Exception(f"❌ Failed to create device flow: {flow.get('error')}")
        
        print(flow['message'])  # Instruct the user to visit the URL and enter the code
        result = self.app.acquire_token_by_device_flow(flow)

        if 'access_token' in result:
            # Save token cache to file
            open(cache_file, 'w').write(self.token_cache.serialize())
            return result['access_token']
    
        raise Exception("Could not acquire access token. Error: {}".format(result.get('error_description', 'No error description available')))

    def check_token_permissions(self, access_token):
        """Check the scopes included in the access token"""
        endpoint = "https://graph.microsoft.com/v1.0/me"
        headers = {'Authorization': f'Bearer {access_token}'}
        
        response = requests.get(endpoint, headers=headers)
        print(response.json())
    
    def test_access_token(self):
        """Test if the access token is valid by fetching OneDrive metadata."""
        if not self.access_token:
            print("❌ No access token available!")
            return
        
        print(f"🔑 Testing Access Token: {self.access_token[:50]}...")  # Print first 50 chars
        test_url = "https://graph.microsoft.com/v1.0/me/drive/root"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(test_url, headers=headers)
        
        if response.status_code == 200:
            print("✅ Access Token is valid!")
        else:
            print(f"❌ Token Test Failed: {response.status_code}, {response.text}")

    
    def refresh_workbook(self, file_path, session_id):
        """Refresh workbook to ensure updates are applied"""
        session_id = self.create_session(file_path)
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/application/calculate"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
            'workbook-session-id': session_id
        }
        data = {"calculationType": "Full"}
        response = self.session.post(endpoint, headers=headers, json=data)
        print(f"Response Status: {response.status_code}")
        print(f"Response Body: {response.text}")
        response.raise_for_status()
        self.close_session(file_path, session_id)
        return response.json()

        
    def create_session(self, file_path):
        """Create a session for updating the Excel file"""
        self.access_token = self.refresh_token()  # Refresh token before creating session
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/createSession"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.post(endpoint, headers=headers, json={"persistChanges": True})
        
        print(f"🔄 Session API Response: {response.status_code}, {response.text}")  # Debug log

        try:
            response.raise_for_status()
        except requests.exceptions.HTTPError as e:
            print(f"❌ Error in create_session: {str(e)}")
            return None

        session_data = response.json()
        print(f"📂 Full Session Response: {session_data}")  # Debugging log

         # Use the complete session ID instead of extracting just the usid
        session_id = session_data.get("id")
        
        if not session_id:
            raise ValueError(f"❌ Failed to get session ID! Response: {session_data}")

        print(f"✅ Session ID: {session_id}")  # Debugging log
        return session_id

    def close_session(self, file_path, session_id):
        """Close the workbook session"""
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/closeSession"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
            'workbook-session-id': session_id
        }
        response = requests.post(endpoint, headers=headers)
        response.raise_for_status()

    def update_excel(self, worksheet_name, range_address, values):
        print(f"📂 Updating {file_path}...")  # Debugging log
    
        try:
            print("⚡ Calling create_session...")
            session_id = self.create_session(file_path)
            print(f"✅ session_id in update: {session_id}")
        except Exception as e:
            print(f"❌ Failed to create session: {str(e)}")
            return None  # Prevent function from proceeding
        
        if not session_id:
            raise ValueError("❌ No session ID returned!")
        
        # Refresh workbook before making changes
        # self.refresh_workbook(file_path, session_id)
        file_path = self._get_user_excel_file()
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_path}:/workbook/worksheets/{worksheet_name}/range(address='{range_address}')"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
            'workbook-session-id': session_id  # Attach session ID
        }
        data = {
            'values': values
        }
        # print(f"Endpoint: {endpoint}")
        # print(f"Payload: {data}")
        response = requests.patch(endpoint, headers=headers, json=data)
        # print(f"Response: {response.status_code}, {response.text}")
        if response.status_code not in [200, 201, 204]:
            print(f"❌ Request failed with status {response.status_code}: {response.text}")
            raise requests.exceptions.RequestException(response)

        print(f"✅ Successfully updated {file_path}")

        # Refresh workbook again after update
        # self.refresh_workbook(file_path, session_id)
        
        self.close_session(file_path, session_id)  # Close session after update
        return response.json()
    
    def _get_user_excel_file(self):
        """Find or create the user's personal Excel file."""
        
        # 📌 Check if the user already has a file
        endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(endpoint, headers=headers).json()

        for item in response.get("value", []):
            if "excel workbook" in item["name"].lower():
                return item["name"]  # Return existing file ID
        
        # ❌ If no file exists, create a copy from a template
        return self._copy_template_excel()

    
    def _copy_template_excel(self):
        """Copy a template Excel file for a new user."""
        
        template_id = "55D6028EF742F77A!s6dcadefacd544780aa6262a995b23109"  # Replace with the actual file ID of the template
        endpoint = f"https://graph.microsoft.com/v1.0/me/drive/items/{template_id}/copy"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        
        data = {
            "parentReference": {"path": "/drive/root:"},
            "name": "Copy of Excel Workbook V1.2.xlsx"
        }

        response = requests.post(endpoint, headers=headers, json=data)
        response.raise_for_status()

        return response.json().get("id")  # Return new file ID

    def read_excel(self, worksheet_name, range_address):
        file_path = self._get_user_excel_file()
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
