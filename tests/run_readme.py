

def run_simple_example():
    import xl_tables as xl
    import datetime
    import time

    xc = xl.Excel()
    xc1 = xl.Excel()

    class MyTable(xl.Table):
        # Constants initialize their value on Table creation
        label_first = xl.Constant('First Name', (1, 1))  # , sheet='Sheet2')
        first_name = xl.Cell((1, 2))
        label_last = xl.Constant('Last Name', (2, 1))
        last_name = xl.Cell(2, 2)
        label_now = xl.Constant('Now', (3, 1))
        now = xl.DateTime(3, 2)
        label_today = xl.Constant('Today', (4, 1))
        today = xl.Date(4, 2)
        label_time = xl.Constant('Time', (5, 1))
        time = xl.Time(5, 2)

        header = xl.Constant(['Data 1', 'Data 2', 'Data 3'], rows=7, row_length=3)
        array_item = xl.RangeItem('A8:C10')  # Contiguous Range is preferable
        array = xl.Range('A8:C10')
        # array_item = xl.RowItem(8, 9, 10, row_length=3)
        # array = xl.Row(8, 9, 10, row_length=3)

    tbl = MyTable(DisplayAlerts=False)

    tbl.Name = 'Book1'
    tbl.first_name = 'John'
    tbl.last_name = 'Doe'
    tbl.now = datetime.datetime.now()
    tbl.today = datetime.datetime.today()
    tbl.time = datetime.time(20, 1, 1)  # datetime.datetime.now().time()

    tbl.array = [(1, 2, 3),
                 (4, 5, 6),
                 (7, 8, 9)]

    # Make a border around the cells in the table
    tbl.array_item.Borders(xl.xlEdgeTop).LineStyle = xl.xlDouble

    text = '{lbl1} = {opt1}\n' \
           '{lbl2} = {opt2}\n' \
           '{lbl3} = {now}\n' \
           '{lbl4} = {today}\n' \
           '{lbl5} = {time}\n' \
           '\n' \
           '{header}\n' \
           '{arr}\n'.format(lbl1=tbl.label_first, opt1=tbl.first_name, lbl2=tbl.label_last, opt2=tbl.last_name,
                            lbl3=tbl.label_now, now=tbl.now, lbl4=tbl.label_today, today=tbl.today,
                            lbl5=tbl.label_time, time=tbl.time,
                            header=tbl.get_row_text(tbl.header, delimiter=', '),
                            arr=tbl.get_table_text(tbl.array, delimiter=', '))

    with open('person_text.txt', 'w') as f:
        f.write(text)
    print('===== Manual Text =====')
    print(text)
    print('===== End =====')

    # Short function provided for this
    txt = tbl.get_table_text(tbl.array, header=tbl.header, head={tbl.label_first: tbl.first_name,
                                                                 tbl.label_last: tbl.last_name,
                                                                 tbl.label_now: tbl.now,
                                                                 tbl.label_today: tbl.today,
                                                                 tbl.label_time: tbl.time})
    print('===== Get Table Text =====')
    print(txt)
    print('===== End =====')

    tbl.save('person.txt')  # 'person.txt' or '.tsv' will save every cell separated by '\t'
    tbl.save('person.csv')  # 'person.csv' will save every cell separated by ','
    tbl.save('person.xlsx')


if __name__ == '__main__':
    run_simple_example()
