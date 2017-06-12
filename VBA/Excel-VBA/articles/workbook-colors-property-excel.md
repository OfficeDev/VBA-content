---
title: Workbook.Colors Property (Excel)
keywords: vbaxl10.chm199088
f1_keywords:
- vbaxl10.chm199088
ms.prod: excel
api_name:
- Excel.Workbook.Colors
ms.assetid: 60fc038b-980b-c1bc-6d1c-69d9d31a11ba
ms.date: 06/08/2017
---


# Workbook.Colors Property (Excel)

Returns or sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value. Read/write  **Variant** .


## Syntax

 _expression_ . **Colors**( **_Index_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The color number (from 1 to 56). If this argument isn?t specified, this method returns an array that contains all 56 of the colors in the palette.|

## Example

This example sets the color palette for the active workbook to be the same as the palette for Book2.xls.


```vb
ActiveWorkbook.Colors = Workbooks("BOOK2.XLS").Colors
```

This example sets color five in the color palette for the active workbook.




```vb
ActiveWorkbook.Colors(5) = RGB(255, 0, 0)
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

