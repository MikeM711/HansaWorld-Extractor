print("\n")

mylist = ['GP100 24x24', 'GP100 27x27', 'GP100 56x56', '518EL']

new_list = []

for x in range(0,4):
    mylist_new_var = mylist[x].split("x",1)[-1]
    x = x + 1
    testing = new_list.append(mylist_new_var)
#    print(mylist_new_var)
print(new_list, "my new list")

