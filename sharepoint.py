import time
import json
import base64
import requests


# TODO: Remove the verify: False from the requests
class Sharepoint:
    def __init__(self, access_token, sp_url):
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        # Initialize site_id and drive_id
        self.drive_id = None
        self.item_id = None
        self.init_ids(sp_url)

    # Set the Item ID & Drive ID for the SharePoint URL
    def init_ids(self, sp_url):
        encoded_url = base64.b64encode(sp_url.encode("utf-8")).decode("utf-8")
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

    # Copy the item to a new location in SharePoint
    def copy(self, dest_drive_id, dest_parent_id, dest_name, item_id=None):
        if item_id is None:
            item_id = self.item_id

        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/copy"
        data = {
            "parentReference": {"driveId": dest_drive_id, "id": dest_parent_id},
            "name": dest_name,
        }
        response = requests.post(api, headers=self.headers, json=data, verify=False)
        if response.status_code == 202:
            location = response.headers.get("Location")
            return self.monitor_copy(location)
        else:
            raise Exception(
                f"Failed to copy the item: {response.status_code} {response.text}"
            )

    # Monitor the copy operation
    def monitor_copy(self, location):
        while True:
            response = requests.get(
                location,
                headers={
                    "Content-Type": "application/json",
                },
                verify=False,
            )
            if response.status_code >= 400:
                raise Exception(
                    f"Failed to monitor the copy operation: {response.status_code} {response.text}"
                )

            result = response.json()
            status = result.get("status")
            if status == "completed":
                print("Copy operation completed successfully.")
                return result.get("resourceId")
            elif status == "failed":
                raise Exception(f"Copy operation failed: {result.get('error')}")
            else:
                # Wait for 5 seconds
                time.sleep(5)

    # Get the file information from the SharePoint
    def get_file_details(self, item_id):
        file_details = []
        items = self.get_children(item_id)
        for item in items:
            if "folder" not in item:
                # Only include files, not folders
                file_name = item["name"]
                file_size = item.get("size", 0)
                file_size_mb = f"{file_size / (1024 * 1024):.2f}MB"
                file_details.append((file_name, file_size_mb))
            else:
                # Recursively collect details from subfolders
                file_details.extend(self.get_file_details(item["id"]))
        return file_details

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

    # Read the data from the excel file
    def excel_read(self, item_id, sheet_name, start_row):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/workbook/worksheets('{sheet_name}')/usedRange"
        response = requests.get(api, headers=self.headers)
        if response.status_code >= 400:
            raise Exception(
                f"Failed to read the excel file: {response.status_code} {response.text}"
            )

        # Get the rows from the start_row
        rows = response.json()
        rows = rows["values"][start_row - 2 :] if rows["values"] is not None else []
        # Filter the non-empty rows
        filtered_rows = []
        for idx, row in enumerate(rows, start=start_row):
            if any(cell for cell in row):
                filtered_rows.append((idx, row))
        return filtered_rows

    # Write the data to the excel file
    def excel_write_row(self, item_id, sheet_name, row_idx, col_start, col_end, values):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/workbook/worksheets('{sheet_name}')/range(address='{col_start}{row_idx}:{col_end}{row_idx}')"
        data = {"values": values}
        response = requests.patch(api, headers=self.headers, data=json.dumps(data))
        if response.status_code >= 400:
            raise Exception(
                f"Failed to write to the excel file: {response.status_code} {response.text}"
            )
