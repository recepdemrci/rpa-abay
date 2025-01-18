import logging


class Model:
    def __init__(self, row):
        # Project detail
        self.oem = row[0]
        self.project = row[1]
        self.system = row[2]
        self.partname = row[3]
        self.partno = row[4]
        # Link to share with supplier
        self.url = row[7]
        # Responsible - Company
        self.r = row[9]
        self.r_email = row[10]
        self.r_cc_email = row[11].split(";") if row[11] else []
        # Mail detail
        self.subject = row[15]
        self.comment = row[16]
        # Responsible - Supplier
        self.sp = row[14]
        self.sp_r = row[17]
        self.sp_r_email = row[18]
        # Flag for send
        self.send = row[19]
        # Share detail
        self.share_url = row[20]
        self.share_date = row[21]
        self.share_status = row[22]
        # Error detail
        self.error = None
        self.valid = self.validate()

    # Validate required fields
    def validate(self):
        # Skip validation if the row is already sent
        # Skip validation if the row not marked for send
        if self.send != "Gönder." or self.share_status == "Gönderildi.":
            return False

        # Check for missing required fields and set error message
        errors = []
        if not self.oem:
            errors.append("OEM")
        if not self.project:
            errors.append("Project")
        if not self.system:
            errors.append("System")
        if not self.partname:
            errors.append("Part Name")
        if not self.partno:
            errors.append("Customer Part Number")
        if not self.url:
            errors.append("Data Link")
        if not self.r:
            errors.append("(Requested By) Responsible")
        if not self.r_email:
            errors.append("(Requested By) Responsible E-mail Address")
        if not self.sp:
            errors.append("Company")
        if not self.sp_r:
            errors.append("(Received By) Responsible")
        if not self.sp_r_email:
            errors.append("(Received By) Responsible E-mail Address")
        if not self.subject:
            errors.append("Subject")
        if not self.comment:
            errors.append("Comment")
        if errors:
            self.error = "Missing Data: " + ", ".join(errors)
            return False

        # Return self for method chaining
        return True
