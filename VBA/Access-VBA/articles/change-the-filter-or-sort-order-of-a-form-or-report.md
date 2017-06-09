---
title: Change the Filter or Sort Order of a Form or Report
ms.prod: access
ms.assetid: 9888dbcd-7409-f334-115e-a318131ebca4
ms.date: 06/08/2017
---


# Change the Filter or Sort Order of a Form or Report

After a form or report is open, you can change the filter or sort order in response to users' actions by setting form and report properties in Visual Basic for Applications (VBA) code. For example, you may want to provide a button or a shortcut menu that users can use to change the records that are displayed. Or you may include an option group control on a form that users can use to select from common sorting options.

To set the filter of a form or report, set its  **Filter** property to the appropriate _wherecondition_ argument, and then set the **FilterOn** property to **True**. To set the sort order, set the **OrderBy** property to the field or fields you want to sort on, and then set the **OrderByOn** property to **True**. If a filter or sort order is already applied on a form, you can change it simply by setting the **Filter** or **OrderBy** properties.

When you apply or change the filter or sort order by setting these properties, Access automatically requeries the records in the form or report. For example, the following code changes the sort order of a form based on a user's selection in an option group:




```vb
Private Sub SortOptionGrp_AfterUpdate() 
 
 Const conName = 0 
 Const conDate = 1 
 
On Error GoTo ErrorHandler 
 
 Select Case SortOptionGrp 
 Case conName 
 Me.OrderBy = "LastName, FirstName" ' Sort by two fields. 
 Case conDate 
 Me.OrderBy = "HireDate DESC" ' Sort by descending date. 
 End Select 
 
 Me.OrderByOn = True ' Apply the sort order. 
 
 Exit Sub 
 
ErrorHandler: 
 MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```

Whether the filter and sort order get set in code or by the user, you can apply or remove them by setting the  **FilterOn** and **OrderByOn** properties to **True** or **False**.

