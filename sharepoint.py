import time
import json
import base64
import logging
import requests


class Sharepoint:
    def __init__(self, access_token, sp_url, verify=True):
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        self.verify = verify
        # Initialize site_id and drive_id
        self.drive_id = None
        self.item_id = None
        self.init_ids(sp_url)

    # Set the Item ID & Drive ID for the SharePoint URL
    def init_ids(self, sp_url):
        encoded_url = base64.b64encode(sp_url.encode("utf-8")).decode("utf-8")
        encoded_url = encoded_url.replace("/", "_").replace("+", "-").replace("=", "")

        api = f"{self.base_url}/shares/u!{encoded_url}/driveItem"
        response = requests.get(api, headers=self.headers, verify=self.verify)
        if response.status_code >= 400:
            logging.error(
                f"Init ItemID & DriveID failed. {response.status_code} {response.text}"
            )
            raise Exception("SharePoint Data Link is invalid.")

        data = response.json()
        self.drive_id = data["parentReference"]["driveId"]
        self.item_id = data["id"]

    # Extract the item ID from the SharePoint URL
    def get_item_id(self):
        if self.item_id is None:
            raise Exception("Item ID is not initialized.")
        return self.item_id

    # Get the list of items (files and folders) from the SharePoint
    def get_children(self, item_id=None):
        if item_id is None:
            item_id = self.item_id

        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/children"
        response = requests.get(api, headers=self.headers, verify=self.verify)
        if response.status_code >= 400:
            logging.error(
                f"Get children failed. {response.status_code} {response.text}"
            )
            raise Exception("Failed to get children from SharePoint.")

        items = response.json().get("value", [])
        return items

    # Find the item ID from a given directory in SharePoint
    def find_dir(self, parent_id, dir_name):
        items = self.get_children(parent_id)
        for item in items:
            if item.get("folder") and item["name"] == dir_name:
                return item["id"]
        logging.error(f"Directory not found: {dir_name} in SharePoint.")
        raise Exception(f"Company directory {dir_name} NOT found in SharePoint.")

    # Copy the item to a new location in SharePoint
    def copy(self, dest_drive_id, dest_parent_id, company, dest_name, item_id=None):
        if item_id is None:
            item_id = self.item_id

        # Search for the destination directory in parent directory in SharePoint
        dest_id = self.find_dir(dest_parent_id, company)

        # Copy the item to the destination directory
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/copy"
        data = {
            "parentReference": {"driveId": dest_drive_id, "id": dest_id},
            "name": dest_name,
        }
        response = requests.post(
            api, headers=self.headers, json=data, verify=self.verify
        )
        if response.status_code == 202:
            location = response.headers.get("Location")
            return self.monitor_copy(location)
        else:
            logging.error(f"Copy failed. {response.status_code} {response.text}")
            raise Exception("SharePoint copy operation failed.")

    # Monitor the copy operation
    def monitor_copy(self, location):
        while True:
            response = requests.get(
                location,
                headers={
                    "Content-Type": "application/json",
                },
                verify=self.verify,
            )
            if response.status_code >= 400:
                logging.error(
                    f"Monitor copy failed. {response.status_code} {response.text}"
                )
                raise Exception("SharePoint copy operation failed.")

            result = response.json()
            status = result.get("status")
            if status == "completed":
                logging.info("Copy operation successful.")
                return result.get("resourceId")
            elif status == "failed":
                logging.error(f"Copy operation failed: {result.get('error')}")
                raise Exception("SharePoint copy operation failed.")
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
        # Create a share link for the item
        share_api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/createLink"
        data = {
            "type": "view",
            "scope": "users",
        }
        response = requests.post(
            share_api, headers=self.headers, json=data, verify=self.verify
        )
        if response.status_code >= 400:
            logging.error(
                f"Create share link failed. {response.status_code} {response.text}"
            )
            raise Exception("Failed to create share link in SharePoint.")
        share_url = response.json()["link"]["webUrl"]

        # Grant access to a specific user by inviting them
        invite_api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/invite"
        data = {
            "message": "You have been granted access to the shared folder.",
            "requireSignIn": True,
            "sendInvitation": False,
            "roles": ["read"],
            "recipients": [{"email": email} for email in emails],
        }
        response = requests.post(
            invite_api, headers=self.headers, json=data, verify=self.verify
        )
        if response.status_code >= 400:
            logging.error(f"Invite user failed. {response.status_code} {response.text}")
            raise Exception("Failed to create share link in SharePoint.")

        logging.info(f"Share link created successfully: {share_url}")
        return share_url

    # Send an email
    def send_email(self, data, dest_name, files):
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
                    <p>Aşağıdaki linkte, {data.sp} Firması için Farplas A.Ş tarafından {dest_name} dosyası erişiminize açılmıştır.</p>
                    <p><b>OEM:</b> {data.oem}</p>
                    <p><b>Project:</b> {data.project}</p>
                    <p><b>System:</b> {data.system}</p>
                    <p><b>Part Name:</b> {data.partname}</p>
                    <p><b>Part Number:</b> {data.partno}</p>
                    <p><b>Link:</b><br><a href="{data.share_url}">{dest_name}</a></p>
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
        response = requests.post(
            api, headers=self.headers, json=message, verify=self.verify
        )
        if response.status_code >= 400:
            logging.error(f"Send email failed. {response.status_code} {response.text}")
            raise Exception("Failed to send email.")
        logging.info(f"Send email to {data.sp_r_email} successfully.")

    # Read the data from the excel file
    def excel_read(self, item_id, sheet_name, start_row):
        api = f"{self.base_url}/drives/{self.drive_id}/items/{item_id}/workbook/worksheets('{sheet_name}')/usedRange"
        response = requests.get(api, headers=self.headers)
        if response.status_code >= 400:
            logging.error(
                f"Read excel file failed. {response.status_code} {response.text}"
            )
            raise Exception("Failed to read the excel file from SharePoint.")

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
            logging.error(
                f"Write excel file failed. {response.status_code} {response.text}"
            )
            raise Exception("Failed to write the excel file to SharePoint.")
