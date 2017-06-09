---
title: Range.CopyFromRecordset Method (Excel)
keywords: vbaxl10.chm144105
f1_keywords:
- vbaxl10.chm144105
ms.prod: excel
api_name:
- Excel.Range.CopyFromRecordset
ms.assetid: cec7fded-f4e0-1b1c-5374-8a860828c9cc
ms.date: 06/08/2017
---


# Range.CopyFromRecordset Method (Excel)

Copies the contents of an ADO or DAO  **Recordset** object onto a worksheet, beginning at the upper-left corner of the specified range. If the **Recordset** object contains fields with OLE objects in them, this method fails.


## Syntax

 _expression_ . **CopyFromRecordset**( **_Data_** , **_MaxRows_** , **_MaxColumns_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Data_|Required| **Variant**|The  **Recordset** object to copy into the range.|
| _MaxRows_|Optional| **Variant**|The maximum number of records to copy onto the worksheet. If this argument is omitted, all the records in the  **Recordset** object are copied.|
| _MaxColumns_|Optional| **Variant**|The maximum number of fields to copy onto the worksheet. If this argument is omitted, all the fields in the  **Recordset** object are copied.|

### Return Value

Long


## Remarks

Copying begins at the current row of the  **Recordset** object. After copying is completed, the **EOF** property of the **Recordset** object is **True** .


## Example

This example copies the field names from a DAO  **Recordset** object into the first row of a worksheet and formats the names as bold. The example then copies the recordset onto the worksheet, beginning at cell A2.


```vb
For iCols = 0 to rs.Fields.Count - 1 
 ws.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name 
Next 
ws.Range(ws.Cells(1, 1), _ 
 ws.Cells(1, rs.Fields.Count)).Font.Bold = True 
ws.Range("A2").CopyFromRecordset rs
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

