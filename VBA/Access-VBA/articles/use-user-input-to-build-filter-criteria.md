---
title: Use User Input to Build Filter Criteria
ms.prod: access
ms.assetid: 0ce3417e-3527-ded4-0940-691c5c81352c
ms.date: 06/08/2017
---


# Use User Input to Build Filter Criteria

The [BuildCriteria](application-buildcriteria-method-access.md) method enables you to easily construct criteria for a filter based on user input. It parses the expression argument in the same way that the expression would be parsed had it been entered in the query design grid, in Filter By Form or Server Filter By Form mode.

The following example prompts the user to enter the first few letters of a product's name and then uses the  **BuildCriteria** method to construct a criteria string based on the user's input. Next, the procedure provides this string as an argument to the **Filter** property of a form named Products. Finally, the **FilterOn** property is set to apply the filter.



```vb
Sub SetFilter() 
    Dim frm As Form, strMsg As String 
    Dim strInput As String, strFilter As String 
 
    ' Open Products form in Form view. 
    DoCmd.OpenForm "Products" 
 
    ' Return Form object variable pointing to Products form. 
    Set frm = Forms!Products 
 
    strMsg = "Enter one or more letters of product name " _ 
        &; "followed by an asterisk." 
 
    ' Prompt user for input. 
    strInput = InputBox(strMsg) 
 
    ' Build criteria string. 
    strFilter = BuildCriteria("ProductName", dbText, strInput) 
 
    ' Set Filter property to apply filter. 
    frm.Filter = strFilter 
 
    ' Set FilterOn property; form now shows filtered records. 
    frm.FilterOn = True 
End Sub
```


