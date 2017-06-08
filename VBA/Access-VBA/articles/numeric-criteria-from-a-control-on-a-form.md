---
title: Numeric Criteria from a Control on a Form
keywords: vbaac10.chm5187731
f1_keywords:
- vbaac10.chm5187731
ms.prod: access
ms.assetid: faa0806c-f349-09a2-1535-582264434233
ms.date: 06/08/2017
---


# Numeric Criteria from a Control on a Form

If you want to change the  _criteria_ argument for an operation based on a user's decision, you can specify that the criteria comes from a control on a form. For example, you could specify that the _criteria_ argument comes from a text box containing the EmployeeID number.

To specify numeric criteria coming from a control on a form, you can include in the  _criteria_ argument an expression that references that control. This control expression should be separate from the string expression, so that Access will evaluate control expression first and concatenate it with the rest of the string expression before performing the appropriate operation.

Suppose that you are performing the [DLookup](application-dlookup-method-access.md) function on an Employees table to find the last name of an employee based on the EmployeeID number. In the following example, the criteria is determined by the current value of the EmployeeID control on the Orders form. The expression referencing the control is evaluated each time the function is called, so that if the value of the control changes, the criteria argument will reflect that change.




```
=DLookup("[LastName]", "Employees", "[EmployeeID] = " _ 
 &; Forms!Orders!EmployeeID)
```

If the current value of the EmployeeID field is 7, the  _criteria_ argument that is passed to the **DLookup** function is:



```
"[EmployeeID] = 7"
```


 **Tip**   To troubleshoot an expression in the _criteria_ argument, break the expression into smaller components and test each individually in the Immediate window. When all of the components are working correctly, put them back together one at a time until the complete expression works correctly.

You can also include a variable representing a numeric value in the  _criteria_ argument. The variable should be separate from the string expression, so that Access will evaluate the variable first and then concatenate it with the rest of the string expression.
The following example shows how to construct a  _criteria_ argument that includes a variable:



```vb
Dim intNum As Integer 
Dim varResult As Variant 
 
intNum = 7 
varResult = DLookup("[LastName]", "Employees", _ 
 "[EmployeeID] = " &; intNum)
```


