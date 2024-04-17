import openpyxl


class ExcelChecker:
    @staticmethod
    def excels_are_equivalent(file1_path: str, file2_path: str):
        # Load the workbooks
        wb1 = openpyxl.load_workbook(file1_path, data_only=True)
        wb2 = openpyxl.load_workbook(file2_path, data_only=True)

        # Check if the number of sheets is the same
        if len(wb1.sheetnames) != len(wb2.sheetnames):
            print("Number of sheets is different")
            return False

        # Create dictionaries to store sheet data
        wb1_sheets = {sheet: wb1[sheet].values for sheet in wb1.sheetnames}
        wb2_sheets = {sheet: wb2[sheet].values for sheet in wb2.sheetnames}

        # Check if the sheet names are the same
        if set(wb1_sheets.keys()) != set(wb2_sheets.keys()):
            print("Sheet names are different")
            return False

        # Compare the values in each sheet
        for sheet in wb1_sheets.keys():
            wb1_values = list(wb1_sheets[sheet])
            wb2_values = list(wb2_sheets[sheet])

            if wb1_values != wb2_values:
                print(f"Values in sheet '{sheet}' are different")
                print(f"wb1: {wb1_values}")
                print(f"wb2: {wb2_values}")
                return False

        # If all checks pass, the workbooks are equivalent
        return True
