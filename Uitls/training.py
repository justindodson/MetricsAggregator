import openpyxl
from utils import build_sheet_from_book

class Training:
    def __init__(self, *args, **kwargs):
        self.sites = []

    def __training_completion_status(self, file_path):
        ws = build_sheet_from_book(file_path)
        