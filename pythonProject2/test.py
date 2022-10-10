import random

tim = []
n = random.randint(5,8)
for i in range(1, n+1):
    l = random.sample(range(1, 21), 4)
    tim.append(l)
print("tim is ",tim)


abs_stu = random.sample(range(1, 91), n)
print("abs_stu is ",abs_stu)
abStu = {absstu:time for absstu,time in zip(abs_stu,tim)}
print(abStu)




