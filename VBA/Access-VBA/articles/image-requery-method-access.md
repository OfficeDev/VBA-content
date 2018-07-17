---
title: Image.Requery Method (Access)
keywords: vbaac10.chm10359
f1_keywords:
- vbaac10.chm10359
ms.prod: access
api_name:
- Access.Image.Requery
ms.assetid: 98f16a2d-ad18-c576-11e0-43d43fcf8859
ms.date: 06/08/2017
---


# Image.Requery Method (Access)

The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.


## Syntax

 _expression_. **Requery**

 _expression_ A variable that represents an **Image** object.


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


 **Note**  


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


[Image Object](image-object-access.md)

