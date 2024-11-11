class Model:
    def __init__(self, sheet, row_idx):
        # Project detail
        self.oem = sheet[f"B{row_idx}"].value
        self.project = sheet[f"C{row_idx}"].value
        self.system = sheet[f"D{row_idx}"].value
        self.partname = sheet[f"E{row_idx}"].value
        self.partno = sheet[f"F{row_idx}"].value
        # Link to share with supplier
        self.url = (
            sheet[f"I{row_idx}"].hyperlink.target
            if sheet[f"I{row_idx}"].hyperlink
            else sheet[f"I{row_idx}"].value
        )
        # Responsible - Farplas
        self.r = sheet[f"K{row_idx}"].value
        self.r_email = sheet[f"L{row_idx}"].value
        self.r_cc_email = (
            sheet[f"M{row_idx}"].value.split(";") if sheet[f"M{row_idx}"].value else []
        )
        # Mail detail
        self.subject = sheet[f"Q{row_idx}"].value
        self.comment = sheet[f"R{row_idx}"].value
        # Responsible - Supplier
        self.sp = sheet[f"P{row_idx}"].value
        self.sp_r = sheet[f"S{row_idx}"].value
        self.sp_r_email = sheet[f"T{row_idx}"].value
        # Flag for send
        self.send = sheet[f"U{row_idx}"].value
        # Share that will be filled by script
        self.share_url = sheet[f"V{row_idx}"].value
        self.share_date = sheet[f"W{row_idx}"].value
        self.share_status = sheet[f"X{row_idx}"].value
