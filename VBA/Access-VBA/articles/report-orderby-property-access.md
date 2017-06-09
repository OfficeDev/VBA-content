---
title: Report.OrderBy Property (Access)
keywords: vbaac10.chm13692
f1_keywords:
- vbaac10.chm13692
ms.prod: access
api_name:
- Access.Report.OrderBy
ms.assetid: 1939157c-12ad-2e58-bf4c-22c04a6c4366
ms.date: 06/08/2017
---


# Report.OrderBy Property (Access)

You can use the  **OrderBy** property to specify how you want to sort records in a report. Read/write **String**.


## Syntax

 _expression_. **OrderBy**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **OrderBy** property is a string expression that is the name of the field or fields on which you want to sort records. When you use more than one field name, separate the names with a comma (,). Use the **OrderBy** property to save an ordering value and apply it at a later time. **OrderBy** values are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they aren't automatically applied.

When you set the  **OrderBy** property by entering one or more field names, the records are sorted in ascending order. Similarly, Visual Basic sorts these fields in ascending order by default.

If you want to sort records in descending order, type DESC at the end of the string expression. For example, to sort customer records in descending order by contact name, set the  **OrderBy** property to "ContactName DESC".

Select the field by which you want to sort the records and either click the appropriate Sort button on the toolbar, or point to  **Sort** on the **Records** menu and click the appropriate command on the submenu. You can also set the **OrderByOn** property for either forms or reports by using Visual Basic.


 **Note**  When a new object is created, it inherits the  **RecordSource**, **Filter**, **OrderBy**, and **OrderByOn** properties of the table or query it was created from. For forms and reports, inherited filters aren't automatically applied when an object is opened.


## See also


#### Concepts


[Report Object](report-object-access.md)

