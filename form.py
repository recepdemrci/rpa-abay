import logging

from model import Model
from sharepoint import Sharepoint


class Form:
    def __init__(self, access_token, spdir_parent, excel_name):
        self.access_token = access_token
        self.excel_name = excel_name
        self.sp_parent = Sharepoint(access_token, spdir_parent, verify=False)
        self.sp_item_id = self.get_item_id()
        self.rows = []

    # Get the item id of the excel file in the parent directory in SharePoint
    def get_item_id(self):
        try:
            items = self.sp_parent.get_children()
            for item in items:
                if item["name"] == self.excel_name:
                    return item["id"]
            # Raise an error if the excel file is not found
            logging.error(
                f"Excel file '{self.excel_name}' not found in the SharePoint folder."
            )
            raise FileNotFoundError(
                f"Excel file '{self.excel_name}' not found in the SharePoint folder."
            )
        except Exception as e:
            logging.error(f"Get Item ID of {self.excel_name} failed. {e}")
            raise

    # Read the data from the excel file in SharePoint
    def read(self, sheet_name):
        try:
            rows = self.sp_parent.excel_read(self.sp_item_id, sheet_name, start_row=6)
            logging.info(f"Read '{self.excel_name}' successfully.")
            
            for idx, item in rows:
                row = Model(item)
                if row.valid or row.error:
                    self.rows.append((idx, row))
            return self.rows
        except Exception as e:
            logging.error(f"Read '{self.excel_name}' failed. {e}")
            raise

    # Write the updated rows to the excel file in SharePoint
    def write(self, sheet_name):
        try:
            for idx, row in self.rows:
                values = [[row.share_url, row.share_date, row.share_status, row.error]]
                self.sp_parent.excel_write_row(
                    self.sp_item_id,
                    sheet_name,
                    row_idx=idx,
                    col_start="V",
                    col_end="Y",
                    values=values,
                )
            logging.info(f"Write to '{self.excel_name}' successfully.")
        except Exception as e:
            logging.error(f"Write to '{self.excel_name}' failed. {e}")
            raise
