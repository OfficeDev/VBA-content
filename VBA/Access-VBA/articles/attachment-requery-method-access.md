---
title: Attachment.Requery Method (Access)
keywords: vbaac10.chm13908
f1_keywords:
- vbaac10.chm13908
ms.prod: access
api_name:
- Access.Attachment.Requery
ms.assetid: 6af04ea8-02cb-9eda-439d-6c69cd772891
ms.date: 06/08/2017
---


# Attachment.Requery Method (Access)

The  **Requery** method updates the data underlying a specified control that's on the active form by requerying the source of data for the control.


## Syntax

 _expression_. **Requery**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

You can use this method to ensure that a form or control displays the most recent data.

The  **Requery** method does one of the following:


- Reruns the query on which the form or control is based.
    
- Displays any new or changed records, or removes deleted records from the table on which the form or control is based.
    
- Updates records displayed based on any changes to the  **Filter** property of the form.
    
Controls based on a query or table include:


- List boxes and combo boxes.
    
- Subform controls.
    
- OLE objects, such as charts.
    
- Controls for which the  **ControlSource** property setting includes domain aggregate functions or SQL aggregate functions.
    
If you specify any other type of control for the object specified by expression, the record source for the form is requeried.

If the object specified by expression isn't bound to a field in a table or query, the  **Requery** method forces a recalculation of the control.

If you omit the object specified by expression, the  **Requery** method requeries the underlying data source for the form or control that has the focus. If the control that has the focus has a record source or row source, it will be requeried; otherwise, the control's data will simply be refreshed.

If a subform control has the focus, this method requeries the record source only for the subform, not the parent form.


 **Note**  


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

