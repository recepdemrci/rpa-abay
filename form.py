from model import Model
from sharepoint import Sharepoint


class Form:
    def __init__(self, access_token, spdir_parent, excel_name):
        self.access_token = access_token
        self.excel_name = excel_name
        self.sp_parent = Sharepoint(access_token, spdir_parent)
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
            raise FileNotFoundError(
                f"Excel file '{self.excel_name}' not found in the SharePoint folder."
            )
        except Exception as e:
            print(f"Error getting item id of '{self.excel_name}': {e}")
            raise

    # Read the data from the excel file in SharePoint
    def read(self, sheet_name):
        try:
            rows = self.sp_parent.excel_read(self.sp_item_id, sheet_name, start_row=6)
            for idx, item in rows:
                row = Model(item)
                if row.send == "Gönder." and not row.share_status == "Gönderildi.":
                    self.rows.append((idx, row))
            print(f"Read '{self.excel_name}' successfully.")
            return self.rows
        except Exception as e:
            print(f"Error reading '{self.excel_name}': {e}")
            raise

    # Write the updated rows to the excel file in SharePoint
    def write(self, sheet_name):
        try:
            for idx, row in self.rows:
                if row.share_status == "Gönderildi.":
                    values = [[row.share_url, row.share_date, row.share_status]]
                    self.sp_parent.excel_write_row(
                        self.sp_item_id,
                        sheet_name,
                        row_idx=idx,
                        col_start="V",
                        col_end="X",
                        values=values,
                    )
            print(f"Updated '{self.excel_name}' successfully.")
        except Exception as e:
            print(f"Error writing to '{self.excel_name}': {e}")
            raise
