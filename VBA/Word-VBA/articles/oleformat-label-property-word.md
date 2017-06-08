---
title: OLEFormat.Label Property (Word)
keywords: vbawd10.chm154337292
f1_keywords:
- vbawd10.chm154337292
ms.prod: word
api_name:
- Word.OLEFormat.Label
ms.assetid: 3603bdee-3259-9068-9dfc-6861c253df97
ms.date: 06/08/2017
---


# OLEFormat.Label Property (Word)

Returns a string that's used to identify the portion of the source file that's being linked. Read-only  **String** .


## Syntax

 _expression_ . **Label**

 _expression_ An expression that returns an **[OLEFormat](oleformat-object-word.md)** object.


## Remarks

If the source file is a Microsoft Excel workbook, the  **Label** property might return "Workbook1!R3C1:R4C2" if the OLE object contains only a few cells from the worksheet.

This property works only for shapes, inline shapes, or fields that are linked OLE objects.


## Example

This example returns the label for the first field in the active document.


```vb
MsgBox ActiveDocument.Fields(1).OLEFormat.Label
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

