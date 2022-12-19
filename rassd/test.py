import itertools

list1 = [i for i in range(1,5)]
list2 = [i for i in range(6,11)]

for (i, j) in itertools.zip_longest(range(10, 44), range(27, 76, 6)):
    print(i, j)