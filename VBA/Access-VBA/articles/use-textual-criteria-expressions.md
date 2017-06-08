---
title: Use Textual Criteria Expressions
ms.prod: access
ms.assetid: 72ee596d-b08c-6af4-041a-6771ac8ce524
ms.date: 06/08/2017
---


# Use Textual Criteria Expressions

To specify textual criteria for an operation, you supply a text string as part of the string expression that forms the  _criteria_ argument. This text string must be enclosed in single quotation marks (').


 **Note**  The single quotation marks indicate to Access that the  _criteria_ argument contains a string within a string.


Suppose that you are using the ADO  **[Find](http://msdn.microsoft.com/library/A7CC9CEB-FDB9-73E2-8328-70B174F93CDA%28Office.15%29.aspx)** method to find the first occurrence of a last name in an Employees table. You could construct the _criteria_ argument as in the following example, which moves the current record pointer to the first record in which an employee's last name is Buchanan. Note that the string literal `Buchanan` is enclosed in single quotation marks and the entire string comprising the criteria argument must also be enclosed in double quotation marks (").




```vb
Dim rst As New ADODB.Connection 
 
rst.open "Employees", CurrentProject.Connection,_ 
     dbOpenDynaset, adlockoptimistic) 
rst.Find "[LastName] = 'Buchanan'"
```


