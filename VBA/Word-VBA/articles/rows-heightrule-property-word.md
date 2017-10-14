---
title: Rows.HeightRule Property (Word)
keywords: vbawd10.chm155975688
f1_keywords:
- vbawd10.chm155975688
ms.prod: word
api_name:
- Word.Rows.HeightRule
ms.assetid: 478635fd-fcaa-d679-e0e2-b24258615d04
ms.date: 06/08/2017
---


# Rows.HeightRule Property (Word)

Returns or sets the rule for determining the height of the specified cells or rows. Read/write  **WdRowHeightRule** .


## Syntax

 _expression_ . **HeightRule**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Example

This example sets the height rule for the selected rows to automatically adjust to the tallest cell in the row.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Rows.HeightRule = wdRowHeightAuto 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

