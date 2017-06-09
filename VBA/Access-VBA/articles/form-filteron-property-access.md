---
title: Form.FilterOn Property (Access)
keywords: vbaac10.chm13347
f1_keywords:
- vbaac10.chm13347
ms.prod: access
api_name:
- Access.Form.FilterOn
ms.assetid: 6ff59ffc-844b-24fc-925f-0331cfcf01ec
ms.date: 06/08/2017
---


# Form.FilterOn Property (Access)

You can use the  **FilterOn** property to specify or determine whether the **Filter** property for a form or report is applied. Read/write **Boolean**.


## Syntax

 _expression_. **FilterOn**

 _expression_ A variable that represents a **Form** object.


## Remarks

If you want to specify a server filter within a Microsoft Access project (.adp) for data located on a server, use the  **ServerFilter** property.

To apply a saved filter, press the  **Apply Filter** button for forms, or apply the filter by using a macro or Visual Basic by setting the **FilterOn** property to **True** for forms or reports. For reports, you can set the **FilterOn** property to Yes in the report's property sheet.

The  **Apply Filter** button indicates the state of the **Filter** and **FilterOn** properties. The button remains disabled until there is a filter to apply. If an existing filter is currently applied, the **Apply Filter** button appears pressed in. To apply a filter automatically when a form or report is opened, specify in the **OnOpen** event property setting of the form either a macro that uses the ApplyFilter action or an event procedure that uses the **ApplyFilter** method of the **DoCmd** object.

You can remove a filter by clicking the pressed-in  **Apply Filter** button, clicking **Remove Filter/Sort** on the **Records** menu, or by using Visual Basic to set the **FilterOn** property to **False**. For reports, you can remove a filter by setting the **FilterOn** property to No in the report's property sheet.


 **Note**  When a new object is created, it inherits the  **RecordSource**, **Filter**, **ServerFilter**. **OrderBy**, and **OrderByOn** properties of the table or query it was created from. For forms and reports, inherited filters aren't automatically applied when an object is opened.


## See also


#### Concepts


[Form Object](form-object-access.md)

