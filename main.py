import os
import time
from datetime import datetime
from dotenv import load_dotenv

from sharepoint import Sharepoint
from form import Form
from auth import get_access_token

# Define the path to your Excel file from environment variables
load_dotenv()
FREQUENCY = int(os.getenv("FREQUENCY"))
EXCEL_NAME = os.getenv("RPABAY_DATA_MANAGEMENT_REQUEST_FORM")
SHEET_NAME = os.getenv("RPABAY_DATA_MANAGEMENT_REQUEST_FORM_SHEET")
SPDIR_BASE = os.getenv("RPABAY_DATA_MANAGEMENT")
SPDIR_SENT = os.getenv("RPABAY_DATA_GIDEN")


def main(access_token):
    try:
        # STEP 1: Download & Read excel file from SharePoint
        request_form = Form(
            access_token, spdir_parent=SPDIR_BASE, excel_name=EXCEL_NAME
        )
        request_form.read(SHEET_NAME)

        # TODO: Implement the loop error handling
        for index, (row_idx, row) in enumerate(request_form.rows):
            print("--------------------------------------------------")
            print(f"Started: {row.sp}, {row.sp_r_email}")

            # STEP 2: Set the destination name on SharePoint
            timestamp = datetime.now().strftime("%Y%m%d%H%M")
            dest_name = f"{row.sp}-{timestamp}"

            # STEP 3: Initialize the source & destination Sharepoint objects
            sp_src = Sharepoint(access_token, row.url)
            sp_dest = Sharepoint(access_token, SPDIR_SENT)

            # STEP 4: Copy the source folder to the destination folder in SharePoint
            dest_item_id = sp_src.copy(sp_dest.drive_id, sp_dest.item_id, dest_name)

            # STEP 5: Share the destination Sharepoint link with the supplier responsible
            share_url = sp_dest.share(
                dest_item_id, [row.sp_r_email, row.r_email, *row.r_cc_email]
            )

            # STEP 6: Send mail to the supplier responsible
            files = sp_dest.get_file_details(dest_item_id)
            sp_dest.send_email(row, dest_name, files)

            # STEP 7: Update the data list with the share link
            row.share_url = share_url
            row.share_date = datetime.now().strftime("%d.%m.%Y")
            row.share_status = "Gönderildi."
            request_form.rows[index] = (row_idx, row)

            print(f"Completed: {row.sp}, {row.sp_r_email}")
            print("--------------------------------------------------")

        # STEP 7: Write & Upload the updated excel file to SharePoint
        request_form.write(SHEET_NAME)

    except Exception as e:
        print(f"[ERROR]: {e}")


if __name__ == "__main__":
    # Authentication
    access_token = get_access_token()
    print("Authenticated successfully")

    while True:
        main(access_token)
        print(f"Sleeping for {FREQUENCY} seconds...")
        time.sleep(FREQUENCY)
