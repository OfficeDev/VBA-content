---
title: DoCmd.SetFilter Method (Access)
keywords: vbaac10.chm6000
f1_keywords:
- vbaac10.chm6000
ms.prod: access
api_name:
- Access.DoCmd.SetFilter
ms.assetid: 98c3e202-8581-2215-7fb2-4a006a97d38f
ms.date: 06/08/2017
---


# DoCmd.SetFilter Method (Access)

Use the  **SetFilter** method to apply a filter to the records in the active datasheet, form, report, or table.


## Syntax

 _expression_. **SetFilter**( ** _FilterName_**, ** _WhereCondition_**, ** _ControlName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FilterName_|Optional|**Variant**|If provided, the name of a query or of a filter saved as a query. This argument or the WhereCondition argument is required.|
| _WhereCondition_|Optional|**Variant**|If provided, a SQL WHERE clause that restricts the records in the datasheet, form, report, or table.|
| _ControlName_|Optional|**Variant**|If provided, the name of the control that corresponds to the subform or subreport to be filtered. If empty, the current object is filtered.|

## Remarks

When you run this method, the filter is applied to the table, form, report or datasheet (for example, query result) that is active and has the focus.

 The **Filter** property of the active object is used to save the WhereCondition argument and apply it at a later time. Filters are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they are not automatically applied.

To automatically apply a filter when the object is opened, set the  **FilterOnLoad** property to True.


## Example

The following code example filters the active object so that it displays only records that begin with "NWTB".


```vb
DoCmd.SetFilter WhereCondition:="[Product Code] Like ""NWTB*"""
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

