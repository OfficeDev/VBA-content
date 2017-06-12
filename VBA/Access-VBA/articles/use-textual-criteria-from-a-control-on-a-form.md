---
title: Use Textual Criteria from a Control on a Form
ms.prod: access
ms.assetid: 236e57eb-3523-92da-e665-281961dfb431
ms.date: 06/08/2017
---


# Use Textual Criteria from a Control on a Form

If you want to change the  _criteria_ argument for an operation based on a user's decision, you can specify that the criteria come from a control on a form. For example, you could specify that the _criteria_ argument comes from a list box containing the last names of all employees in an Employees table.

To specify textual criteria coming from a control on a form, you include in the  _criteria_ argument an expression that references the control on the form. This expression should be separate from the string expression, so that Access will evaluate the control expression first and concatenate it with the rest of the string expression before performing the appropriate operation.

In addition to enclosing the entire string expression in double quotation marks ("), you must also ensure that the textual criteria within the string expression are enclosed in single quotation marks ('). The quotation marks must be included in the strings flanking the expression that references the control on the form.


 **Note**  The single quotation marks indicate to Access that the  _criteria_ argument contains a string within a string.

The following example performs a lookup on an Employees table and returns the region in which an employee lives, based on the employee's last name. The current value of a list box control called LastName on the Employees form determines the criteria. Note the placement of the single quotation marks. 



```
=DLookup("[Region]", "Employees", "[LastName] = '" _
    &; Forms!Employees!LastName &; "'")
```

If the current value of the control is  `King`, the following  _criteria_ argument is passed to the [DLookup](application-dlookup-method-access.md) function after Access evaluates the expression and concatenates the strings:



```
"[LastName] = 'King'"

```

Keep in mind that the entire string comprising the  _criteria_ argument must also be enclosed in double quotation marks once the strings have been concatenated.
To troubleshoot an expression in the  _criteria_ argument, break the expression into smaller components and test each one individually in the Immediate window. When all of the components are working correctly, put them back together one at a time until the complete expression works correctly.
You can also include a variable representing a textual string in the  _criteria_ argument. The variable should be separate from the string expression, so that Access will evaluate the variable first and then concatenate it with the rest of the string expression. The textual string must be enclosed in single or double quotation marks.
The following example shows how to construct a  _criteria_ argument that includes a variable representing a textual string:



```vb
Dim strLastName As String
Dim varResult As Variant

strLastName = "King"
varResult = DLookup("[EmployeeID]", "Employees", "[LastName] = '" _
    &; strLastName &; "'")
```


