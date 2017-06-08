---
title: TextBox.Requery Method (Access)
keywords: vbaac10.chm11034
f1_keywords:
- vbaac10.chm11034
ms.prod: access
api_name:
- Access.TextBox.Requery
ms.assetid: b1f8991e-7ccc-4f0b-c50f-1d51a0abda7e
ms.date: 06/08/2017
---


# TextBox.Requery Method (Access)

The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.


## Syntax

 _expression_. **Requery**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

You can use this method to ensure that a form or control displays the most recent data.

The  **Requery** method does one of the following:


- Reruns the query on which the form or control is based.
    
- Displays any new or changed records or removes deleted records from the table on which the form or control is based.
    
- Updates records displayed based on any changes to the  **Filter** property of the form.
    
Controls based on a query or table include:


- List boxes and combo boxes.
    
- Subform controls.
    
- OLE objects, such as charts .
    
- Controls for which the  **ControlSource** property setting includes domain aggregate functions or SQL aggregate function.
    
If you specify any other type of control for the object specified by expression, the record source for the form is requeried.

If the object specified by expression isn't bound to a field in a table or query, the  **Requery** method forces a recalculation of the control.

If you omit the object specified by expression, the  **Requery** method requeries the underlying data source for the form or control that has the focus. If the control that has the focus has a record source or row source, it will be requeried; otherwise, the control's data will simply be refreshed.

If a subform control has the focus, this method only requeries the record source for the subform, not the parent form.


|**Note**|
|:-----|
|<ul><li>The **Requery** method updates the data underlying a form or control to reflect records that are new to or deleted from the record source since it was last queried. The **Refresh** method shows only changes that have been made to the current set of records; it doesn't reflect new or deleted records in the record source. The **Repaint** method simply repaints the specified form and its controls.</li><li>The **Requery** method doesn't pass control to the operating system to allow Windows to continue processing messages. Use the **DoEvents** function if you need to relinquish temporary control to the operating system.</li><li>The **Requery** method is faster than the **Requery** action. When you use the **Requery** action, Microsoft Access closes the query and reloads it from the database. When you use the **Requery** method, Microsoft Access reruns the query without closing and reloading it.</li></ul>|  


## Example

The following example uses the  **Requery** method to requery the data from the EmployeeList list box on an Employees form:


```vb
Public Sub RequeryList() 
 
    Dim ctlCombo As Control 
 
    ' Return Control object pointing to a combo box. 
    Set ctlCombo = Forms!Employees!ReportsTo 
 
    ' Requery source of data for list box. 
    ctlCombo.Requery 
 
End Sub
```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

