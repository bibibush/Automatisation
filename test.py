import math

y = 7 * (0.3  * 93) + (0.3 * 93) - 202.2
x=  y / 93
x = round(x, 1)
z = math.ceil(x)
if (z * 10) - (x * 10) >= 5:
    print(z - 0.5)
else:
    print(z)