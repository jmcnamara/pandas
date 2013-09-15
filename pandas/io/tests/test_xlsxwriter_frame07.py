###############################################################################
#
# Tests for Pandas ExcelWriter xlsxwriter option.
#

import unittest
import os
from pandas.core.api import DataFrame
import pandas.util.testing as testutil
from pandas.io.excel import ExcelWriter
from .xlsxwriter_test_helper import _compare_xlsx_files


class TestCompareXLSXFiles(unittest.TestCase):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.maxDiff = None

        filename = 'xw_frame07.xlsx'
        test_dir = testutil.get_data_path()
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_to_excel(self):
        """Test the creation of a simple workbook using to_excel()."""
        filename = self.got_filename

        ####################################################

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': [2, 4, 6, 8]})

        df.to_excel(filename,
                    sheet_name='Sheet1',
                    startcol=2,
                    startrow=1,
                    header=False,
                    index=False,
                    engine='xlsxwriter')

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_excelwriter_to_excel(self):
        """Test the creation of a simple workbook using ExcelWriter()."""
        filename = self.got_filename

        ####################################################

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': [2, 4, 6, 8]})

        writer = ExcelWriter(filename, engine='xlsxwriter')

        df.to_excel(writer,
                    sheet_name='Sheet1',
                    startcol=2,
                    startrow=1,
                    header=False,
                    index=False)

        writer.save()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_to_excel_with_config(self):
        """Test workbook creation using to_excel() and pandas.config."""
        filename = self.got_filename

        ####################################################

        from pandas import set_option
        set_option('io.excel.xlsx.writer', 'xlsxwriter')

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': [2, 4, 6, 8]})

        df.to_excel(filename,
                    sheet_name='Sheet1',
                    startcol=2,
                    startrow=1,
                    header=False,
                    index=False)

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_excelwriter_to_excel_with_config(self):
        """Test workbook creation using ExcelWriter() and pandas.config."""
        filename = self.got_filename

        ####################################################

        from pandas import set_option
        set_option('io.excel.xlsx.writer', 'xlsxwriter')

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': [2, 4, 6, 8]})

        writer = ExcelWriter(filename)

        df.to_excel(writer,
                    sheet_name='Sheet1',
                    startcol=2,
                    startrow=1,
                    header=False,
                    index=False)

        writer.save()

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def tearDown(self):
        # Cleanup.
        if os.path.exists(self.got_filename):
            os.remove(self.got_filename)

if __name__ == '__main__':
    unittest.main()
