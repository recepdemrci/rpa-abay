import os
import base64
import requests
from urllib.parse import urlparse, unquote


# TODO: Remove the verify: False from the requests
class Sharepoint:
    def __init__(self, access_token, sp_url):
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        # Initialize site_id and drive_id
        self.sp_url = sp_url
        # hostname, path = self.parse_sharepoint_url()
        self.drive_id = None
        self.item_id = None
        self.init_ids()

    # Extract the hostname and path from the SharePoint URL
    def parse_sharepoint_url(self):
        # Parse the URL to extract components
        parsed_url = urlparse(self.sp_url)
        # Hostname (e.g., 'bilisimpark-my.sharepoint.com')
        hostname = parsed_url.netloc
        # Extract the folder path after '/s/' (which is the site path)
        # Decode the URL path (SharePoint encodes special characters)
        # TODO: Handle more complex SharePoint URLs
        path = parsed_url.path.split("/s/")[1]
        path = path.split("/")[0]
        path = unquote(path)
        return hostname, path

    # Set the Item ID & Drive ID for the SharePoint URL
    def init_ids(self):
        encoded_url = base64.b64encode(self.sp_url.encode("utf-8")).decode("utf-8")
        encoded_url = encoded_url.replace("/", "_").replace("+", "-").replace("=", "")

        api = f"{self.base_url}/shares/u!{encoded_url}/driveItem"
        response = requests.get(api, headers=self.headers, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to set Item ID & Drive ID: {response.status_code} {response.text}"
            )

        data = response.json()
        self.drive_id = data["parentReference"]["driveId"]
        self.item_id = data["id"]

    # Extract the item ID from the SharePoint URL
    def get_item_id(self):
        if self.item_id is None:
            raise Exception("Item ID is not set.")
        return self.item_id

    # Get the list of items (files and folders) from the SharePoint
    def get_children(self, item_id=None):
        if item_id is None:
            item_id = self.item_id

        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/children"
        response = requests.get(api, headers=self.headers, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to get the list of files: {response.status_code} {response.text}"
            )

        items = response.json().get("value", [])
        return items

    # Lock the item
    def checkout(self, item_id):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/checkout"
        response = requests.post(api, headers=self.headers, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to checkout the item: {response.status_code} {response.text}"
            )

    # Checkin the item
    def checkin(self, item_id):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/checkin"
        data = {"comment": "RPABAY: Dosya paylaşımı tamamlandı."}
        response = requests.post(api, headers=self.headers, json=data, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to checkin the item: {response.status_code} {response.text}"
            )

    # Discard the checkout
    def discard_checkout(self, item_id):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/discardCheckout"
        requests.post(api, headers=self.headers, verify=False)

    # Download a file from the sharepoint
    def download_file(self, dest_path, item_id=None):
        if item_id is None:
            item_id = self.item_id

        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/content"
        response = requests.get(api, headers=self.headers, stream=True, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to download the file: {response.status_code} {response.text}"
            )
        # Save to destination path
        with open(dest_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)

    # Download all files from a sharepoint folder, including nested folders
    def download(self, dest, item_id=None):
        if item_id is None:
            item_id = self.item_id

        # Get the list of items (files and folders) from the SharePoint or OneDrive URL
        items = self.get_children(item_id)

        # Loop through the items in the sharepoint folder
        #         If it's a folder, create a corresponding directory and recurse
        #         If it's a file, download it
        details = []
        for item in items:
            item_id = item["id"]
            item_name = item["name"]
            item_size = item.get("size")
            item_type = item.get("folder")

            if item_type:
                # Recursively call the function to download files from this subfolder
                sub_dest_dir = os.path.join(dest, item_name)
                os.makedirs(sub_dest_dir, exist_ok=True)
                self.download(sub_dest_dir, item_id)
            else:
                # Download the file
                dest_path = os.path.join(dest, item_name)
                self.download_file(dest_path, item_id)

            # Append the file name and size to the details list
            item_size_mb = (
                f"{item_size / (1024 * 1024):.2f}MB" if item_size is not None else "-"
            )
            details.append((item_name, item_size_mb))

        # Return file names and sizes
        print(f"Downloaded files to '{dest}' successfully.")
        return details

    # Upload a file to the sharepoint
    def upload_file(self, src, item_id=None):
        if item_id is None:
            item_id = self.item_id

        file_name = os.path.basename(src)
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}:/{file_name}:/content?@microsoft.graph.conflictBehavior=replace"
        with open(src, "rb") as file_data:
            response = requests.put(
                api, headers=self.headers, data=file_data, verify=False
            )
            if response.status_code >= 400:
                raise Exception(
                    f"Failed to upload file: {response.status_code}: {response.text}"
                )

    # Recursivelly upload the folder to a SharePoint url
    def upload(self, src, item_id=None):
        if item_id is None:
            item_id = self.item_id

        # Create subfolder in SharePoint for each directory
        folder_name = os.path.basename(src)
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/children"
        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        }
        response = requests.post(api, headers=self.headers, json=data, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to create folder {folder_name}: {response.status_code} {response.text}"
            )
        # Get the ID of the newly created folder
        new_item_id = response.json().get("id")

        # Get a list of files and directories in the current directory
        items = os.listdir(src)
        for item in items:
            item_path = os.path.join(src, item)
            # Upload files into sharepoint
            if os.path.isfile(item_path):
                self.upload_file(item_path, new_item_id)
            # Recursively upload subfolders
            if os.path.isdir(item_path):
                self.upload(item_path, new_item_id)

        print(f"Uploaded '{src}' successfully.")
        return new_item_id

    # Create a share link and give permission
    # for given item_id to given emails in SharePoint
    def share(self, item_id, emails):
        # Step 1: Create a share link for the item
        share_api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/createLink"
        data = {
            "type": "view",
            "scope": "users",
        }
        response = requests.post(
            share_api, headers=self.headers, json=data, verify=False
        )
        if response.status_code >= 400:
            raise Exception(
                f"Failed to create share link: {response.status_code} {response.text}"
            )
        share_url = response.json()["link"]["webUrl"]

        # Step 2: Grant access to a specific user by inviting them
        invite_api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/invite"
        data = {
            "message": "You have been granted access to the shared folder.",
            "requireSignIn": True,
            "sendInvitation": False,
            "roles": ["read"],
            "recipients": [{"email": email} for email in emails],
        }
        response = requests.post(
            invite_api, headers=self.headers, json=data, verify=False
        )
        if response.status_code >= 400:
            raise Exception(
                f"Failed to invite user: {response.status_code} {response.text}"
            )

        print(f"Created share link {share_url} successfully.")
        return share_url

    # Send an email
    def send_email(self, data, dirname, files):
        # Generate the file listing with names and sizes
        file_list = ""
        for idx, (name, size) in enumerate(files, start=1):
            file_list += f"{idx}-{name}<br>Dosya boyutu: {size}<br>"

        # Message content
        message = {
            "message": {
                "subject": f"FARPLAS_RPABAY_DATA_PAYLASIMI_{data.sp}_{data.subject}",
                "body": {
                    "contentType": "HTML",
                    "content": f"""
                    <p>Merhaba {data.sp_r},</p>
                    <p>Aşağıdaki linkte, {data.sp} Firması için Farplas A.Ş tarafından {dirname} dosyası erişiminize açılmıştır.</p>
                    <p><b>OEM:</b> {data.oem}</p>
                    <p><b>Project:</b> {data.project}</p>
                    <p><b>System:</b> {data.system}</p>
                    <p><b>Part Name:</b> {data.partname}</p>
                    <p><b>Part Number:</b> {data.partno}</p>
                    <p><b>Link:</b><br><a href="{data.share_url}">{dirname}</a></p>
                    <p><b>Dosya içeriği:</b><br>{file_list}</p>
                    <p><b>Yorum / Talep:</b></p>
                    <p><b>{data.comment}</b></p>
                    <p><b>Farplas Sorumlusu:</b> {data.r}</p>
                    <p>İyi çalışmalar.</p>
                    """,
                },
                "toRecipients": [{"emailAddress": {"address": data.sp_r_email}}],
                "ccRecipients": [{"emailAddress": {"address": data.r_email}}]
                + [{"emailAddress": {"address": cc}} for cc in data.r_cc_email],
            }
        }
        # Send the email
        api = f"{self.base_url}/me/sendMail"
        response = requests.post(api, headers=self.headers, json=message, verify=False)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to send email: {response.status_code}, {response.text}"
            )
        print(f"Sent email {data.sp_r_email} successfully.")
