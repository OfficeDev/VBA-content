---
title: ListRows.Add Method (Excel)
keywords: vbaxl10.chm740078
f1_keywords:
- vbaxl10.chm740078
ms.prod: excel
api_name:
- Excel.ListRows.Add
ms.assetid: 32213e09-fd25-3787-3ab8-45ee1249ca1c
ms.date: 06/08/2017
---


# ListRows.Add Method (Excel)

Adds a new row to the table represented by the specified [ListObject](listobject-object-excel.md).


## Syntax

 _expression_ . **Add**( **_Position_** , **_AlwaysInsert_** )

 _expression_ A variable that represents a **ListRows** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Position_|Optional| **Variant**| **Integer** . Specifies the relative position of the new row.|
| _AlwaysInsert_|Optional| **Variant**| **Boolean** . Specifies whether to always shift data in cells below the last row of the table when the new row is inserted, regardless if the row below the table is empty. If **True** , the cells below the table will be shifted down one row. If **False** , if the row below the table is empty, the table will expand to occupy that row without shifting cells below it; but if the row below the table contains data, those cells will be shifted down when the new row is inserted.|

### Return Value

A [ListRow](listrow-object-excel.md) object that represents the new row.


## Remarks

If  _Position_ is not specified, a new bottom row is added. If _AlwaysInsert_ is not specified, the cells below the table will be shifted down one row (same as specifying **True** ).


## Example

The following example adds a new row to the default  **ListObject** object in the first worksheet of the workbook. Because no position is specified, the new row is added to the bottom of the list.


```vb
Set myNewRow = ActiveWorkbook.Worksheets(1).ListObject(1).ListRows.Add
```


## See also


#### Concepts


[ListRows Object](listrows-object-excel.md)

