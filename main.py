import os
import time
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

from sharepoint import Sharepoint
from form import Form
from auth import get_access_token

# Define the path to your Excel file from environment variables
load_dotenv()
FREQUENCY = int(os.getenv("FREQUENCY"))
EXCEL_NAME = os.getenv("RPABAY_DATA_MANAGEMENT_REQUEST_FORM")
SPDIR_BASE = os.getenv("RPABAY_DATA_MANAGEMENT")
SPDIR_SENT = os.getenv("RPABAY_DATA_GIDEN")
LDIR_SENT = os.getenv("RPABAY_DATA_GIDEN_LOCAL")


# Create a folder with the current timestamp
def create_folder(dest_dir, prefix):
    timestamp = datetime.now().strftime("%Y%m%d%H%M")
    dest_path = os.path.join(dest_dir, f"{prefix}-{timestamp}")
    os.makedirs(dest_path, exist_ok=True)
    print(f"Created folder {prefix}-{timestamp} successfully.")
    return dest_path


def main(access_token):
    try:
        # STEP 1: Download & Read excel file from SharePoint
        request_form = Form(
            access_token, spdir_parent=SPDIR_BASE, excel_name=EXCEL_NAME
        )
        request_form.lock()
        request_form.download()
        request_form.read()

        # TODO: Implement the loop error handling
        for index, (row_idx, row) in enumerate(request_form.rows):
            print("--------------------------------------------------")
            print(f"Started: {row.sp}, {row.sp_r_email}")

            # STEP 2: Create a dest directory
            local_dir = create_folder(LDIR_SENT, prefix=row.sp)

            # STEP 3: Download the files from the source Sharepoint
            sp_src = Sharepoint(access_token, row.url)
            item_id = sp_src.get_item_id()
            files = sp_src.download(item_id, local_dir)

            # STEP 4: Upload the files to the destination Sharepoint
            sp_dest = Sharepoint(access_token, SPDIR_SENT)
            item_id = sp_dest.get_item_id()
            child_item_id = sp_dest.upload(item_id, local_dir)

            # STEP 5: Share the destination Sharepoint link with the supplier responsible
            share_url = sp_dest.share(
                child_item_id, [row.sp_r_email, row.r_email, *row.r_cc_email]
            )

            # STEP 6: Send mail to the supplier responsible
            # sp_dest.send_email(row, local_dir.split("\\")[-1], files)

            # STEP 7: Update the data list with the share link
            row.share_url = share_url
            row.share_date = datetime.now().strftime("%d.%m.%Y")
            row.share_status = "Gönderildi."
            request_form.rows[index] = (row_idx, row)

            print(f"Completed: {row.sp}, {row.sp_r_email}")
            print("--------------------------------------------------")

        # STEP 7: Write & Upload the updated excel file to SharePoint
        request_form.write()
        request_form.upload()
        request_form.unlock()

    except Exception as e:
        if request_form:
            request_form.unlock(discard=True)
        print(f"[ERROR]: {e}")


if __name__ == "__main__":
    # Authentication
    access_token = get_access_token()
    print("Authenticated successfully")

    while True:
        main(access_token)
        print(f"Sleeping for {FREQUENCY} seconds...")
        time.sleep(FREQUENCY)
