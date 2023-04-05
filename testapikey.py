# -*- coding: utf-8 -*-
import json

file = open("cypress.json")
A = json.load(file.read())
print(A)

if A == {"a": 1, "b": 2, "c": 3}:
    print("YES")
else:
    print("NO")

