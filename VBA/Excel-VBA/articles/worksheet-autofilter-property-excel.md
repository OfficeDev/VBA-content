---
title: Worksheet.AutoFilter Property (Excel)
keywords: vbaxl10.chm175144
f1_keywords:
- vbaxl10.chm175144
ms.prod: excel
api_name:
- Excel.Worksheet.AutoFilter
ms.assetid: 766f8501-dae7-32a7-9fae-70a87d0a8eba
ms.date: 06/08/2017
---


# Worksheet.AutoFilter Property (Excel)

Returns an  **AutoFilter** object if filtering is on. Read-only.


## Syntax

 _expression_ . **AutoFilter**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

The property returns  **Nothing** if filtering is off.

To create an  **AutoFilter** object for a worksheet, you must turn autofiltering on for a range on the worksheet either manually or using the **AutoFilter** method of the **Range** object.


## Example

The following example returns autofilter for the current worksheet.


```vb
Dim Worksheet1 As Worksheet 
 
Dim returnValue As AutoFilter 
Set returnValue = Worksheet1.AutoFilter
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

