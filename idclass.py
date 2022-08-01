"""
Classes to use for unprocessed document reports
"""


import datetime as dt


##########################
# Class for Document IDs #
##########################

class DocumentID:
    def __init__(self, docID):
        self.docID = str(docID)

    # Parses Transaction ID from first 13 digits
    def get_trid(self):
        tr_id = ("".join(self.docID[0:13]))
        return tr_id

    # Returns last 3 digits
    def get_last(self):
        last = ("".join(self.docID[13:17]))
        return last

    # Parses date from first 6 digits
    @staticmethod
    def get_date(date_str):
        year = int("".join(date_str[0:4]))
        month = int("".join(date_str[4:6]))
        day = int("".join(date_str[6:8]))
        date = dt.date(year, month, day)
        return date


####################
# Class for emails #
####################

class Letters:
    def __init__(self, *args):
        # For not submitted documents (doc ID, amt, date)
        if isinstance(args[2], int):
            self.doc_id = args[0]
            self.amount = "$ " + str(args[1])
            self.year = str(args[2])
        # For rejection documents (doc ID, amt, reasons)
        else:
            self.doc_id = args[0]
            self.amount = "$ " + str(args[1])
            self.reasons = args[2]

    # Formatting
    bold = '\033[1m'
    end = '\033[0m'

    # Makes text bold
    @classmethod
    def bold_text(cls, string):
        return cls.bold + string + cls.end

    # Returns description for rejection documents
    def get_rej_description(self):
        description = self.amount + "\nYour document submission was rejected for the following reason(" \
                                               "s):\n" + self.reasons
        return description

    # Returns description for not submitted documents
    def get_nosub_description(self):
        description = "In " + self.year + " you paid to have documents recorded. Research indicates to date your " \
                            "documents have not been recorded and you have a credit balance of " + self.amount
        return description
