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
        # Responsible - Farplas
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
        # Share that will be filled by script
        self.share_url = row[20]
        self.share_date = row[21]
        self.share_status = row[22]
