import os
import time
import logging
from datetime import datetime
from dotenv import load_dotenv

from sharepoint import Sharepoint
from form import Form
from auth import get_access_token

# Load environment variables from .env file
load_dotenv()
FREQUENCY = int(os.getenv("FREQUENCY"))
EXCEL_NAME = os.getenv("RPABAY_DATA_MANAGEMENT_REQUEST_FORM")
SHEET_NAME = os.getenv("RPABAY_DATA_MANAGEMENT_REQUEST_FORM_SHEET")
SPDIR_BASE = os.getenv("RPABAY_DATA_MANAGEMENT")
SPDIR_SENT = os.getenv("RPABAY_DATA_GIDEN")

# Configure logging
log_file = os.path.join(os.path.dirname(__file__), "app.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
)


def main(access_token):
    try:
        # STEP 1: Read excel file from SharePoint
        request_form = Form(
            access_token, spdir_parent=SPDIR_BASE, excel_name=EXCEL_NAME
        )
        request_form.read(SHEET_NAME)

        for index, (idx, row) in enumerate(request_form.rows):
            try:
                logging.info(f"|-------> ROW {idx} <-------|")

                # Skip the row if there is an error
                if row.error:
                    raise Exception(row.error)

                # STEP 2: Set the destination name on SharePoint
                timestamp = datetime.now().strftime("%Y%m%d%H%M")
                dest_name = f"{row.sp}-{timestamp}"

                # STEP 3: Initialize the source & destination Sharepoint objects
                sp_src = Sharepoint(access_token, row.url, verify=False)
                sp_dest = Sharepoint(access_token, SPDIR_SENT, verify=False)

                # STEP 4: Copy the source folder to the destination folder in SharePoint
                dest_item_id = sp_src.copy(sp_dest.drive_id, sp_dest.item_id, dest_name)

                # STEP 5: Share the destination Sharepoint link with the supplier responsible
                share_url = sp_dest.share(
                    dest_item_id, [row.sp_r_email, row.r_email, *row.r_cc_email]
                )
                row.share_url = share_url

                # STEP 6: Send mail to the supplier responsible
                files = sp_dest.get_file_details(dest_item_id)
                sp_dest.send_email(row, dest_name, files)

                # Write result to row (SUCCESS)
                row.share_status = "Gönderildi."
                row.error = ""
                logging.info(f"COMPLETED ({row.sp} - {row.sp_r_email})")
            except Exception as e:
                # Write result to row (ERROR)
                row.share_status = "Hata."
                row.error = str(e)
                logging.error(f"{e}")
            finally:
                # Write result to row
                row.share_date = datetime.now().strftime("%d.%m.%Y")
                request_form.rows[index] = (idx, row)

        # STEP 7: Write the updated rows to SharePoint
        request_form.write(SHEET_NAME)

    except Exception as e:
        logging.error(f"{e}")


if __name__ == "__main__":
    try:
        # Authentication
        access_token = get_access_token()

        while True:
            start_time = time.time()
            logging.info(
                "=============================START============================="
            )

            main(access_token)

            end_time = time.time()
            elapsed_time = end_time - start_time
            logging.info(f"Elapsed time: {elapsed_time:.2f} seconds")
            logging.info(f"Waiting for {FREQUENCY} seconds...")
            logging.info(
                "==============================END=============================="
            )
            time.sleep(FREQUENCY)

    except Exception as e:
        logging.critical(f"{e}")
