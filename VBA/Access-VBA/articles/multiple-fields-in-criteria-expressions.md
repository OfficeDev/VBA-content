---
title: Multiple Fields in Criteria Expressions
keywords: vbaac10.chm5187501
f1_keywords:
- vbaac10.chm5187501
ms.prod: access
ms.assetid: 6ed0992f-912f-10ca-27de-4e3c9d6a2f71
ms.date: 06/08/2017
---


# Multiple Fields in Criteria Expressions

You can specify multiple fields in a  _criteria_ argument.

To specify multiple fields in the  _criteria_ argument, you must ensure that multiple string expressions are concatenated correctly to form a valid SQL WHERE clause. In an SQL WHERE clause with multiple fields, fields may be joined with one of three keywords: **AND**, **OR**, or **NOT**. Your expression must evaluate to a string that includes one of these keywords.

For example, suppose that you wish to set the  **Filter** property of an Employees form to display records restricted by two sets of criteria. The following example filters the form so that it displays only those employees whose title is "Sales Representative" and who were hired since January 1, 1993:




```vb
Dim datHireDate As Date 
Dim strTitle As String 
 
datHireDate = #1/1/93# 
strTitle = "Sales Representative" 
 
Forms!Employees.Filter = "[HireDate] >= #" &; _ 
 datHireDate &; "# AND [Title] = '" &; strTitle &; "'" 
Forms!Employees.FilterOn = True
```

The  _criteria_ argument evaluates to the following string:



```sql
"[HireDate] >= #1-1-93# AND [Title] = 'Sales Representative'"
```


 **Tip**  To troubleshoot an expression in the  _criteria_ argument, break the expression into smaller components and test each individually in the Immediate window. When all of the components are working correctly, put them back together one at a time until the complete expression works correctly.


