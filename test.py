import unittest
from xlsx_parser import XlsxParaser
from pathlib import Path

class TestStringMethods(unittest.TestCase):
    def test_sheets(self):
        print("TESTING SHEETS")

        xlsx_parser = XlsxParaser(aPath=str(Path(Path.cwd(), "test.xlsx")))
        xlsx_parser.open()

        # Test create sheet
        xlsx_parser.create_sheet("test_sheet")
        self.assertEqual(xlsx_parser.get_sheets(), ['Sheet1', 'test_sheet'])

        # Test rename with valid parameters
        xlsx_parser.rename_sheet("Sheet1", "Sheet2")
        self.assertEqual(xlsx_parser.get_sheets(), ['Sheet2', 'test_sheet'])

        # Test rename with invalid paramenters
        self.assertRaises(RuntimeError, xlsx_parser.rename_sheet, "Sheet1", "Sheet2")

        # Test set sheet with invalid parameters
        self.assertRaises(RuntimeError, xlsx_parser.set_sheet, "Non existing")

        # Test delete with valid parameters
        xlsx_parser.delete_sheet("Sheet2")
        self.assertEqual(xlsx_parser.get_sheets(), ['test_sheet'])

        # Test delete with invalid parameters
        self.assertRaises(RuntimeError, xlsx_parser.delete_sheet, "Sheet3")

    def test_headers(self):
        print("TESTING HEADERS")

        xlsx_parser = XlsxParaser(aPath=str(Path(Path.cwd(), "test.xlsx")))
        xlsx_parser.open()
        
        # Test headers with error in key
        xHeaders = [
            {
                "non_existing": "test_header_1",
                "index": "A",
                "start": 1
            },
            {
                "header": "test_header_2",
                "index": "B",
                "start": 1
            }
        ]
        bStatus = xlsx_parser.set_headers(xHeaders)
        self.assertEqual(bStatus, False)
        self.assertEqual(xlsx_parser.sheet["B1"].value, None)
        
        # Test headers without errors
        xHeaders = [
            {
                "header": "test_header_1",
                "index": "A",
                "start": 1
            },
            {
                "header": "test_header_2",
                "index": "B",
                "start": 1
            }
        ]
        xlsx_parser.set_headers(xHeaders)
        self.assertEqual(xlsx_parser.sheet["A1"].value, "test_header_1")
        self.assertEqual(xlsx_parser.sheet["B1"].value, "test_header_2")

        # Test set headers automatic
        xlsx_parser = XlsxParaser(aPath=str(Path(Path.cwd(), "test.xlsx")))
        xlsx_parser.open()

        xlsx_parser.sheet["A1"] = "test_header_1"
        xlsx_parser.sheet["A2"] = "test_data_1"
        xlsx_parser.sheet["B1"] = "test_header_2"
        xlsx_parser.sheet["B2"] = "test_data_2"
        xlsx_parser.find_headers()
        self.assertEqual(len(xlsx_parser.columns), 2)
        self.assertEqual(xlsx_parser.columns[0].header, "test_header_1")
        self.assertEqual(xlsx_parser.columns[0].start, 1)
        self.assertEqual(xlsx_parser.columns[0].end, 2)

        # Test that the values are present
        self.assertEqual(xlsx_parser.sheet["A1"].value, "test_header_1")
        self.assertEqual(xlsx_parser.sheet["A2"].value, "test_data_1")
        self.assertNotEqual(xlsx_parser.sheet["A3"], "test_data_1")

        # Test get headers
        xHeaders = xlsx_parser.get_headers()
        self.assertEqual(xHeaders, ['test_header_1', 'test_header_2'])

    def test_rows(self):
        print("TESTING ROWS")

        xlsx_parser = XlsxParaser(aPath=str(Path(Path.cwd(), "test.xlsx")))
        xlsx_parser.open()

        xlsx_parser.sheet["A1"] = "NAME"
        xlsx_parser.sheet["B1"] = "SURENAME"
        xlsx_parser.sheet["C1"] = "AGE"

        xlsx_parser.sheet["A2"] = "John"
        xlsx_parser.sheet["B2"] = "Smith"
        xlsx_parser.sheet["C2"] = "36"

        xlsx_parser.sheet["A3"] = "George"
        xlsx_parser.sheet["B3"] = "Simmons"
        xlsx_parser.sheet["C3"] = "48"

        xlsx_parser.sheet["A4"] = "George"
        xlsx_parser.sheet["B4"] = "Hanks"
        xlsx_parser.sheet["C4"] = "52"

        xlsx_parser.sheet["A5"] = "AMY"
        xlsx_parser.sheet["B5"] = "Beerhouse"
        xlsx_parser.sheet["C5"] = "27"

        xlsx_parser.find_headers()

        # Test get row
        xRow = xlsx_parser.get_rows({"row": 2})
        self.assertEqual(xRow, [{
            "row": 2,
            "data": {'NAME': 'John', 'SURENAME': 'Smith', 'AGE': '36'}
        }])

        # Test get empty row
        xRow = xlsx_parser.get_rows({"row": 99})
        self.assertEqual(xRow, [{"row": None, "data": {}}])

        # Test get row by search
        xRow = xlsx_parser.get_rows({"header": "NAME", "search": "George"})
        self.assertEqual(len(xRow), 2)
        self.assertEqual(xRow, [
            {
                "row": 3,
                "data": {'NAME': 'George', 'SURENAME': 'Simmons', 'AGE': '48'}
            },
            {
                "row": 4,
                "data": {'NAME': 'George', 'SURENAME': 'Hanks', 'AGE': '52'}
            }
        ])

        # Test get row by search with non existing header
        xRow = xlsx_parser.get_rows({"header": "CITY", "search": "NEW YORK"})
        self.assertEqual(xRow, [{"row": None, "data": {}}])

        # Test get row by search with non existing search value
        xRow = xlsx_parser.get_rows({"header": "NAME", "search": "ANNA"})
        self.assertEqual(xRow, [{"row": None, "data": {}}])

        # Test get row by search with wrong parameters
        #xRow = xlsx_parser.get_rows({"non_existing": "some data", "search": "some other data"})
        #self.assertEqual(xRow, [{"row": None, "data": {}}])

        xlsx_parser.create_sheet("vehicles")
        xlsx_parser.sheet["A1"] = "AUTOMOBILES"
        xlsx_parser.sheet["B1"] = "MOTORBIKES"
        xlsx_parser.find_headers()

        # Test append row with param bAppend_if_none
        xlsx_parser.append_rows([{"header": "AUTOMOBILES", "data": "BMW M3"}], bAppend_if_none=False)
        self.assertEqual(xlsx_parser.sheet["A2"].value, "BMW M3")

        # Test append row with param bAppend_if_none
        xlsx_parser.append_rows([{"header": "MOTORBIKES", "data": "Hayabusa"}], bAppend_if_none=False)
        self.assertEqual(xlsx_parser.sheet["B2"].value, "Hayabusa")

        # Test append row with partial data without bAppend_if_none=False (you should do this if
        # dont provide all columns data because it will cause fragmentation)
        xlsx_parser.append_rows([{"header": "AUTOMOBILES", "data": "Audi A3"}])
        self.assertEqual(xlsx_parser.sheet["A3"].value, "Audi A3")
        self.assertEqual(xlsx_parser.sheet["B3"].value, None)

        #
        # TODO: add this functionality
        #
        # Test append rows with non existing header
        # bResult = xlsx_parser.append_rows([{"header": "PLANES", "data": "Boeing 737"}])
        # self.assertEqual(bResult, False)

        #
        # NOTE: this is similar to update row ...
        #
        # Test append rows with row parmeter
        xlsx_parser.append_rows([{"header": "AUTOMOBILES", "data": "Ford Escort", "row": 2}], aSheet_name="vehicles")
        self.assertEqual(xlsx_parser.sheet["A2"].value, "Ford Escort")

        # Test update rows with non existing parameter
        bStatus = xlsx_parser.update_rows([{"non_existing": "PLANES", "data": "Boeing 737"}], aSheet_name="vehicles")
        self.assertEqual(bStatus, False)

        # Test update rows 
        xlsx_parser.update_rows([{"header": "AUTOMOBILES", "old_data": "Ford Escort", "new_data": "BMW M4"}], aSheet_name="vehicles")
        self.assertEqual(xlsx_parser.sheet["A2"].value, "BMW M4")

        # Test update rows with insert_if_not_found = True
        xlsx_parser.update_rows([{"header": "AUTOMOBILES", "old_data": "Ferrari 488", "new_data": "Alfa Romeo Giulia"}], aSheet_name="vehicles", insert_if_not_found=True)
        self.assertEqual(xlsx_parser.sheet["A4"].value, "Alfa Romeo Giulia")

        # Test update rows with insert_if_not_found = False
        xlsx_parser.update_rows([{"header": "AUTOMOBILES", "old_data": "Ferrari 488", "new_data": "Ferrari F40"}], aSheet_name="vehicles", insert_if_not_found=False)
        self.assertEqual(xlsx_parser.sheet["A5"].value, None)

        # Test remove rows with wrong argument type
        bResult = xlsx_parser.remove_rows("wrong_type", aSheet_name="vehicles")
        self.assertEqual(bResult, False)

        # Check that the first sheet is still intact
        xlsx_parser.set_sheet("Sheet1")
        self.assertEqual(xlsx_parser.sheet["A2"].value, "John")

        xlsx_parser.remove_rows([2, 4], aSheet_name="Sheet1")
        self.assertEqual(xlsx_parser.sheet["A1"].value, "NAME")
        self.assertEqual(xlsx_parser.sheet["A2"].value, "George")
        self.assertEqual(xlsx_parser.sheet["A3"].value, "AMY")
        self.assertEqual(xlsx_parser.sheet["B2"].value, "Simmons")
        self.assertEqual(xlsx_parser.sheet["B3"].value, "Beerhouse")

        # test append row with sheet (note that we are not currently on vehicles sheet)
        xlsx_parser.append_rows([{"header": "AUTOMOBILES", "data": "Nissan Skyline"}], aSheet_name="vehicles")
        self.assertEqual(xlsx_parser.sheet["A5"].value, "Nissan Skyline")

        # test append row with non existing sheet
        xlsx_parser.append_rows([{"header": "CITY", "data": "London"}], aSheet_name="Sheet1")
        self.assertRaises(RuntimeError, xlsx_parser.append_rows, [{"header": "CITY", "data": "London"}], aSheet_name="countries")

if __name__ == '__main__':
    unittest.main()