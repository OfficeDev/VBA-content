---
title: Filter a Report Using a Form's Filter
ms.prod: access
ms.assetid: 2b029c13-5abd-4865-cd05-25d094a97b9f
ms.date: 06/08/2017
---


# Filter a Report Using a Form's Filter

The following example illustrates how to open a report based on the filtered contents of a form. To do this, specify the form's  **[Filter](form-filter-property-access.md)** property as the value of the **[OpenReport](docmd-openreport-method-access.md)** method's _WhereCondition_ argument.


```vb
Private Sub cmdOpenReport_Click() 
    If Me.Filter = "" Then 
        MsgBox "Apply a filter to the form first." 
    Else 
        DoCmd.OpenReport "rptCustomers", acViewReport, , Me.Filter 
    End If 
End Sub
```


