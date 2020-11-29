import pandas


data = pandas.read_excel('two_lists.xlsx')
list1 = data['list1'].dropna().to_list()
list2 = data['list2'].dropna().to_list()

print("\nNot in list 2:")
for item in list2:
    if item not in list1:
        print(item)

print("\nNot in list 1:")
for item in list1:
    if item not in list2:
        print(item)