---
title: Selection.PasteAndFormat Method (Word)
keywords: vbawd10.chm158663669
f1_keywords:
- vbawd10.chm158663669
ms.prod: word
api_name:
- Word.Selection.PasteAndFormat
ms.assetid: 7ed87209-b786-280e-f3f0-dd81eda6f82d
ms.date: 06/08/2017
---


# Selection.PasteAndFormat Method (Word)

Pastes the selected table cells and formats them as specified.


## Syntax

 _expression_ . **PasteAndFormat**( **_Type_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[WdRecoveryType](wdrecoverytype-enumeration-word.md)**|The type of formatting to use when pasting the selected table cells.|

## Example

This example pastes a selected Microsoft Excel chart as a picture. This example assumes that the Clipboard contains an Excel chart.


```vb
Sub PasteChart() 
 Selection.PasteAndFormat Type:=wdChartPicture 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

