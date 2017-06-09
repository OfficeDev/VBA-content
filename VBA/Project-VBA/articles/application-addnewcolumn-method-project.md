---
title: Application.AddNewColumn Method (Project)
keywords: vbapj.chm710
f1_keywords:
- vbapj.chm710
ms.prod: project-server
api_name:
- Project.Application.AddNewColumn
ms.assetid: 009071ad-b713-4252-ab1c-781d58620d8c
ms.date: 06/08/2017
---


# Application.AddNewColumn Method (Project)

Adds a new column in a specified position, in views where columns can be added.


## Syntax

 _expression_. **AddNewColumn**( ** _Column_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**Variant**|Specifies the absolute column location. A value of 0 adds a column in the left-most position.|

### Return Value

 **Boolean**


## Remarks

If the  _Column_ parameter is omitted, **AddNewColumn** inserts a column to the left of the active column, and displays **[Type Column Name]** in the column heading. By comparison, the[ColumnInsert](application-columninsert-method-project.md) method displays the **Field Settings** dialog box for the new column.


## Example

The following example selects the third column in the current view, and then adds a column to the right of the selected column. In the default  **Gantt Chart** view, the third column is **Task Name**.


```
SelectColumn (2) 
AddNewColumn (3)
```


 **Note**  If the user does not name the column header,  **AddNewColumn** removes the selected column. When you add a column, it does not exist until the field is named. If you try to use the **ColumnEdit** method after **AddNewColumn**, Project shows run-time error 1100 (the command in the macro is not available in this situation).


