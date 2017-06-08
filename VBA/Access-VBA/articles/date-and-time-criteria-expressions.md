---
title: Date and Time Criteria Expressions
keywords: vbaac10.chm10885
f1_keywords:
- vbaac10.chm10885
ms.prod: access
ms.assetid: fff89f87-444e-b275-c7b1-4c82240e57f0
ms.date: 06/08/2017
---


# Date and Time Criteria Expressions

To specify date or time criteria for an operation, you supply a date or time value as part of the string expression that forms the  _criteria_ argument. This value must be enclosed in number signs (#).


 **Note**  The number signs indicate to Access that the  _criteria_ argument contains a date or time within a string.


Suppose that you are creating a filter for an Employees form to display records for all employees born on or after January 1, 1960. You could construct the  _criteria_ argument for the form's **Filter** or **ServerFilter** property as in the following example. Note the placement of the number signs.




```vb
Forms!Employees.Filter = "[BirthDate] >= #1-1-1960#"
```


