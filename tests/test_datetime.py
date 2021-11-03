

def test_datetime():
    import xl_tables as xl
    import datetime
    import time

    class MyTable(xl.Table):
        # Constants initialize their value on Table creation
        now = xl.DateTime(3, 2)
        today = xl.Date(4, 2)
        time = xl.Time(5, 2)
        cus = xl.Cell(6, 2)

    tbl = MyTable('test_datetime.xls')
    print('now', tbl.now)
    print('today', tbl.today)
    print('time', tbl.time)
    print('cus', tbl.cus)

    tbl.now = now = datetime.datetime.now()
    tbl.today = now.date()
    tbl.time = now.time()
    tbl.cus = 0
    tbl.save('test_datetime(now).xls')
    # Using NumberFormat must match an existing format exactly! Change in Excel if you want to see something different.
    # https://docs.microsoft.com/en-us/office/vba/api/excel.range.formatconditions


if __name__ == '__main__':
    test_datetime()

    print('All tests finished successfully!')
