---
title: Table.AllowAutoFit Property (Word)
keywords: vbawd10.chm156303470
f1_keywords:
- vbawd10.chm156303470
ms.prod: word
api_name:
- Word.Table.AllowAutoFit
ms.assetid: e8894734-68b3-60bb-7623-9497e4e99e10
ms.date: 06/08/2017
---


# Table.AllowAutoFit Property (Word)

Allows Microsoft Word to automatically resize cells in a table to fit their contents. Read/write  **Boolean** .


## Syntax

 _expression_ . **AllowAutoFit**

 _expression_ A variable that represents a **[Table](table-object-word.md)** object.


## Example

This example sets the first table in the active document to automatically resize based on its contents.


```vb
Sub AllowFit() 
 ActiveDocument.Tables(1).AllowAutoFit = True 
End Sub
```


## See also


#### Concepts


[Table Object](table-object-word.md)

