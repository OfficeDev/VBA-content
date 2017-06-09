---
title: Report.Filter Property (Access)
keywords: vbaac10.chm13690
f1_keywords:
- vbaac10.chm13690
ms.prod: access
api_name:
- Access.Report.Filter
ms.assetid: ce57e05d-c3a8-012a-205e-7dfb2e4dd78d
ms.date: 06/08/2017
---


# Report.Filter Property (Access)

You can use the  **Filter** property to specify a subset of records to be displayed when a filter is applied to a form, reportquery, or table. Read/write **String**.


## Syntax

 _expression_. **Filter**

 _expression_ A variable that represents a **Report** object.


## Remarks

If you want to specify a server filter within a Microsoft Access project (.adp) for data located on a server, use the  **ServerFilter** property.

The  **Filter** property is a string expression consisting of a WHERE clause without the WHERE keyword. For example, the following Visual Basic code defines and applies a filter to show only customers from the USA:




```vb
Me.Filter = "Country = 'USA'" 
Me.FilterOn = True
```


 **Note**  Setting the  **Filter** property has no effect on the ADO **Filter** property.

You can use the  **Filter** property to save a filter and apply it at a later time. Filters are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they aren't automatically applied.

When a new object is created, it inherits the  **RecordSource**, **Filter**, **OrderBy**, and **OrderByOn** properties of the table or query it was created from.

To apply a saved filter to a form, query, or table, you can click  **Apply Filter** on the toolbar, click **Apply Filter/Sort** on the **Records** menu, or use a macro or Visual Basic to set the **FilterOn** property to **True**. For reports, you can apply a filter by setting the **FilterOn** property to Yes in the report's property sheet.

The  **Apply Filter** button indicates the state of the **Filter** and **FilterOn** properties. The button remains disabled until there is a filter to apply. If an existing filter is currently applied, the **Apply Filter** button appears pressed in.

To apply a filter automatically when a form is opened, specify in the  **OnOpen** event property setting of the form either a macro that uses the ApplyFilter action or an event procedure that uses the **ApplyFilter** method of the **DoCmd** object.

You can remove a filter by clicking the pressed-in  **Apply Filter** button, clicking **Remove Filter/Sort** on the **Records** menu, or using Visual Basic to set the **FilterOn** property to **False**.

When the  **Filter** property is set in form Design view, Microsoft Access does not attempt to validate the SQL expression. If the SQL expression is invalid, an error occurs when the filter is applied.


## See also


#### Concepts


[Report Object](report-object-access.md)

