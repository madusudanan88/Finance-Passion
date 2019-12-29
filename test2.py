from ERD import *
a="Shares"
e1 = erd("E:\Finance Passion\Deposits.xls")
print(a,e1.get_current_value_by_category(a))
a="Savings"
print(a,e1.get_current_value_by_category(a))
a="Others"
print(a,e1.get_current_value_by_category(a))

