import xl_tables as xl


class Person(xl.Table):
    now = xl.DateTime(3, 2)
    today = xl.Date(4, 2)
    time = xl.Time(5, 2)


p = Person('person.csv')
print(p.Cells(1, 1).Value, p.Cells(1, 2).Value)
print(repr(p.now))
print(repr(p.today))
print(repr(p.time))


p.save('new_person.xlsx')
p.save('new_person.txt')
