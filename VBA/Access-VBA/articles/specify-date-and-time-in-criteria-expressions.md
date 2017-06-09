---
title: Specify Date and Time in Criteria Expressions
ms.prod: access
ms.assetid: 749379e7-5fbe-3371-a780-ca7915d8de43
ms.date: 06/08/2017
---


# Specify Date and Time in Criteria Expressions

To specify date or time criteria for an operation, you supply a date or time value as part of the string expression that forms the  _criteria_ argument. This value must be enclosed in number signs (#).


 **Note**  The number signs indicate to Access that the  _criteria_ argument contains a date or time within a string.


Suppose that you are creating a filter for an Employees form to display records for all employees born on or after January 1, 1960. You could construct the  _criteria_ argument for the form's **[Filter](form-filter-property-access.md)** or **[ServerFilter](form-serverfilter-property-access.md)** property, as shown in the following example. Note the placement of the number signs.




```vb
Forms!Employees.Filter = "[BirthDate] >= #1-1-1960#"
```


