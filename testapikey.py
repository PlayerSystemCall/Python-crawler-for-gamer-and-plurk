# -*- coding: utf-8 -*-
import os
import json

file = open("cypress.json")
A = file.read()
print(A)

print(os.getcwd("cypress.json"))

if A == '{"a": 1, "b": 2, "c": 3}':
    print("YES")
else:
    print("NO")

