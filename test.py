amount = [100, 200]
part = [2,3]
amountnew = []
for n, i in enumerate(part):
    for j in range(int(i)):
        amountnew.append(amount[n])

print(amountnew)
print(int((amountnew[1]*int('2')-int('4'))*float('1.05')))