---
title: Worksheet.SetBackgroundPicture Method (Excel)
keywords: vbaxl10.chm175076
f1_keywords:
- vbaxl10.chm175076
ms.prod: excel
api_name:
- Excel.Worksheet.SetBackgroundPicture
ms.assetid: 5cff4730-24ba-6147-76c9-e1f9eb970989
ms.date: 06/08/2017
---


# Worksheet.SetBackgroundPicture Method (Excel)

Sets the background graphic for a worksheet.


## Syntax

 _expression_ . **SetBackgroundPicture**( **_Filename_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the graphic file.|

## Example

This example sets the background graphic for worksheet one.


```vb
Worksheets(1).SetBackgroundPicture "c:\graphics\watermark.gif"
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

