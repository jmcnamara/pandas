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

        filename = 'xw_frame01.xlsx'
        test_dir = testutil.get_data_path()
        self.got_filename = test_dir + '_test_' + filename
        self.exp_filename = test_dir + filename

        self.ignore_files = []
        self.ignore_elements = {}

    def test_to_excel(self):
        """Test the creation of a simple workbook using to_excel."""
        filename = self.got_filename

        ####################################################

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': ['foo', 'bar', 'bin', 'boo']})

        df.to_excel(filename,
                    sheet_name='Sheet1',
                    header=False,
                    index=False,
                    writer_options={'use_xlsxwriter': True})

        ####################################################

        got, exp = _compare_xlsx_files(self.got_filename,
                                       self.exp_filename,
                                       self.ignore_files,
                                       self.ignore_elements)

        self.assertEqual(got, exp)

    def test_excelwriter_to_excel(self):
        """Test the creation of a simple workbook using ExcelWriter."""
        filename = self.got_filename

        ####################################################

        df = DataFrame({'A': [10, 11, 12, 13],
                        'B': ['foo', 'bar', 'bin', 'boo']})

        writer = ExcelWriter(filename,
                             writer_options={'use_xlsxwriter': True})

        df.to_excel(writer,
                    sheet_name='Sheet1',
                    header=False,
                    index=False,
                    writer_options={'use_xlsxwriter': True})

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
