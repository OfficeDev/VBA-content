---
title: Numeric Criteria Expressions
keywords: vbaac10.chm5187729
f1_keywords:
- vbaac10.chm5187729
ms.prod: access
ms.assetid: ff497f13-7251-9131-459f-9bd2b189816b
ms.date: 06/08/2017
---


# Numeric Criteria Expressions

To specify numeric criteria for an operation, you supply a numeric value as part of the string expression that forms the  _criteria_ argument.

Suppose that you are performing the [Application.DLookup Method (Access)](application-dlookup-method-access.md) function on an Employees table to find the last name of a particular employee, and you want to use a value from the EmployeeID field in the function's _criteria_ argument. You could construct a _criteria_ argument like the following example, which returns the last name of the employee whose EmployeeID is 7:



```
=DLookup("[LastName]", "Employees", "[EmployeeID] = 7")
```


