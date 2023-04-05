# -*- coding: utf-8 -*-
file = open("cypress.json")
A = file.read()
print(A)

if A == {"a": 1, "b": 2, "c": 3}:
    print("YES")
else:
    print("NO")

