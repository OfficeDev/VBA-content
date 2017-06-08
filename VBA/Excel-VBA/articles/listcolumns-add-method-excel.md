---
title: ListColumns.Add Method (Excel)
keywords: vbaxl10.chm736073
f1_keywords:
- vbaxl10.chm736073
ms.prod: excel
api_name:
- Excel.ListColumns.Add
ms.assetid: a1127989-f1e0-3c0a-e2c9-24b166c5e001
ms.date: 06/08/2017
---


# ListColumns.Add Method (Excel)

Adds a new column to the list object.


## Syntax

 _expression_ . **Add**( **_Position_** )

 _expression_ A variable that represents a **ListColumns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Optional| **Variant**| **Integer** . Specifies the relative position of the new column that starts at 1. The previous column at this position is shifted outward.|

### Return Value

A  **[ListColumn](listcolumn-object-excel.md)** object that represents the new column.


## Remarks

If  _Position_ is not specified, a new rightmost column is added. A name for the column is automatically generated. The name of the new column can be changed after the column is added.


## Example

The following example adds a new column to the default  **[ListObject](listobject-object-excel.md)** object in the first worksheet of the workbook. Because no position is specified, a new rightmost column is added.


```vb
Set myNewColumn = ActiveWorkbook.Worksheets(1).ListObjects(1).ListColumns.Add
```


## See also


#### Concepts


[ListColumns Object](listcolumns-object-excel.md)

