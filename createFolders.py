import os
import re


class config_folders():
    def __init__(self, today, day, month, year):
        self.today = today
        self.day = day
        self.month = month
        self.year = year
        self.mainFolder = r'\\se10orgfps01\Clearing & Custody Services\1. Product Management\Listing Management\New strikes\\'

    def create_folders(self):

        yearFolder = self.mainFolder + self.year
        if not os.path.exists(yearFolder):
            os.makedirs(yearFolder)

        monthFolder = yearFolder + '\\' + self.month + '\\'
        if not os.path.exists(monthFolder):
            os.makedirs(monthFolder)

        return monthFolder