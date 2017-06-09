---
title: Use Numeric Criteria in Expressions
ms.prod: access
ms.assetid: c2055190-8d65-7342-19ef-582c05846b5b
ms.date: 06/08/2017
---


# Use Numeric Criteria in Expressions

To specify numeric criteria for an operation, you supply a numeric value as part of the string expression that forms the  _criteria_ argument.

Suppose that you are performing the  **[DLookup](application-dlookup-method-access.md)** function on an Employees table to find the last name of a particular employee, and you want to use a value from the EmployeeID field in the function's _criteria_ argument. You could construct a _criteria_ argument like the following example, which returns the last name of the employee whose EmployeeID is 7:



```
=DLookup("[LastName]", "Employees", "[EmployeeID] = 7")
```


