---
title: DoCmd.SetOrderBy Method (Access)
keywords: vbaac10.chm5980
f1_keywords:
- vbaac10.chm5980
ms.prod: access
api_name:
- Access.DoCmd.SetOrderBy
ms.assetid: 020fde6d-4809-79f6-3da5-fc5f6a315a83
ms.date: 06/08/2017
---


# DoCmd.SetOrderBy Method (Access)

Use the  **SetOrderBy** method to apply a sort to the active datasheet, form, report, or table.


## Syntax

 _expression_. **SetOrderBy**( ** _OrderBy_**, ** _ControlName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OrderBy_|Required|**Variant**|A string expression that includes the name of the field or fields on which to sort records and the optional ASC or DESC keywords.|
| _ControlName_|Optional|**Variant**|If provided and the active object is a form or report, the name of the control that corresponds to the subform or subreport that will be sorted. If empty and the active object is a form or report, the parent form or report is sorted.|

## Remarks

When you run this method, the sort is applied to the table, form, report or datasheet (for example, query result) that is active and has the focus. 

The Order By argument is the name of the field or fields on which you want to sort records. When you use more than one field name, separate the names with a comma (,). The  **OrderBy** property of the active object is used to save the ordering value and apply it at a later time. OrderBy values are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they are not automatically applied.

When you set the Order By argument by entering one or more field names and then run the method, the records are sorted by default in ascending order. 

To sort records in descending order, type DESC at the end of the Order By argument expression. For example, to sort customer records in descending order by contact name, set the Order By argument to "ContactName DESC". To sort names by LastName descending, and FirstName ascending, set the Order By argument to "LastName DESC, FirstName ASC" 


## Example

The following code example sorts the active datasheet, form, report or table by LastName descending and FirstName ascending.


```vb
DoCmd.SetOrderBy "LastName DESC, FirstName ASC"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

