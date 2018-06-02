n = int(input())
#fl = 0

for res in range(1, n + 1, 1):
    for fl in range(res):
        print('+___ ' * fl)
        print('|', res, ' / ', sep='')
        print('|__\ ')
        print('|    ')