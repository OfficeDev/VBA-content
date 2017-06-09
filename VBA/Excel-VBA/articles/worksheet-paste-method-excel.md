---
title: Worksheet.Paste Method (Excel)
keywords: vbaxl10.chm175115
f1_keywords:
- vbaxl10.chm175115
ms.prod: excel
api_name:
- Excel.Worksheet.Paste
ms.assetid: 65561666-7a47-29d6-2a5d-b5de642a064f
ms.date: 06/08/2017
---


# Worksheet.Paste Method (Excel)

Pastes the contents of the Clipboard onto the sheet.


## Syntax

 _expression_ . **Paste**( **_Destination_** , **_Link_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Destination_|Optional| **Variant**|A  **Range** object that specifies where the Clipboard contents should be pasted. If this argument is omitted, the current selection is used. This argument can be specified only if the contents of the Clipboard can be pasted into a range. If this argument is specified, the _Link_ argument cannot be used.|
| _Link_|Optional| **Variant**| **True** to establish a link to the source of the pasted data. If this argument is specified, the _Destination_ argument cannot be used. The default value is **False** .|

## Remarks

If you don't specify the  _Destination_ argument, you must select the destination range before you use this method.

This method may modify the sheet selection, depending on the contents of the Clipboard.


## Example

This example copies data from cells C1:C5 on Sheet1 to cells D1:D5 on Sheet1.


```vb
Worksheets("Sheet1").Range("C1:C5").Copy 
ActiveSheet.Paste Destination:=Worksheets("Sheet1").Range("D1:D5")
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

