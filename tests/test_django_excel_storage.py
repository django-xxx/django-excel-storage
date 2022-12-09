# -*- coding: utf-8 -*-

from django_excel_storage import ExcelStorage


class TestDjangoExcelStorageCommands(object):

    def setup_class(self):
        self.data1 = [{
            'Column 1': 1,
            'Column 2': 2,
        }, {
            'Column 1': 3,
            'Column 2': 4,
        }]
        self.data2 = [
            ['Column 1', 'Column 2'],
            [1, 2],
            [3, 4]
        ]
        self.data3 = [
            ['Column 1', 'Column 2'],
            [1, [2, 3]],
            [3, 4]
        ]

        self.data11 = {
            'Sheet 1': [{
                'Column 1': 1,
                'Column 2': 2,
            }, {
                'Column 1': 3,
                'Column 2': 4,
            }]
        }
        self.data22 = {
            'Sheet 1': [
                ['Column 1', 'Column 2'],
                [1, 2],
                [3, 4]
            ]
        }
        self.data33 = {
            'Sheet 1': [
                ['Column 1', 'Column 2'],
                [1, [2, 3]],
                [3, 4]
            ]
        }

        self.content_type_csv = 'text/csv'
        self.content_type_xls = 'application/vnd.ms-excel'

        self.sheet_data1 = {'Sheet 1': [['Column 1', 'Column 2'], [1, 2], [3, 4]]}
        self.sheet_data2 = {'Sheet 1': [['Column 1', 'Column 2'], [1, [2, 3]], [3, 4]]}

    def test_as_csv(self):
        csv1 = ExcelStorage(self.data1, 'my_data', force_csv=True, font='name SimSum')
        assert isinstance(csv1, ExcelStorage)

        csv2 = ExcelStorage(self.data2, 'my_data', force_csv=True, font='name SimSum')
        assert isinstance(csv2, ExcelStorage)

        csv3 = ExcelStorage(self.data3, 'my_data', force_csv=True, font='name SimSum')
        assert isinstance(csv3, ExcelStorage)

    def test_as_xls(self):
        xls1 = ExcelStorage(self.data1, 'my_data', font='name SimSum')
        assert isinstance(xls1, ExcelStorage)

        xls2 = ExcelStorage(self.data2, 'my_data', font='name SimSum')
        assert isinstance(xls2, ExcelStorage)

        # xls3 = ExcelResponse(self.data3, 'my_data', font='name SimSum')
        # assert isinstance(xls3, ExcelStorage)

        xls11 = ExcelStorage(self.data11, 'my_data', font='name SimSum')
        assert isinstance(xls11, ExcelStorage)

        xls22 = ExcelStorage(self.data22, 'my_data', font='name SimSum')
        assert isinstance(xls22, ExcelStorage)

    def test_as_row_merge_xls(self):
        xls1 = ExcelStorage(self.data1, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls1, ExcelStorage)

        xls2 = ExcelStorage(self.data2, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls2, ExcelStorage)

        xls3 = ExcelStorage(self.data3, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls3, ExcelStorage)

        xls11 = ExcelStorage(self.data1, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls11, ExcelStorage)

        xls22 = ExcelStorage(self.data2, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls22, ExcelStorage)

        xls33 = ExcelStorage(self.data3, 'my_data', font='name SimSum', row_merge=True)
        assert isinstance(xls33, ExcelStorage)
