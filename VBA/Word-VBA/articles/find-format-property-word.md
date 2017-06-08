---
title: Find.Format Property (Word)
keywords: vbawd10.chm162529308
f1_keywords:
- vbawd10.chm162529308
ms.prod: word
api_name:
- Word.Find.Format
ms.assetid: 999041b0-e1eb-8155-405b-62475cb57f9d
ms.date: 06/08/2017
---


# Find.Format Property (Word)

 **True** if formatting is included in the find operation. Read/write **Boolean** .


## Syntax

 _expression_ . **Format**

 _expression_ Required. A variable that represents a **[Find](find-object-word.md)** object.


## Example

This example removes all bold formatting in the active document.


```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Format = True 
 .Replacement.ClearFormatting 
 .Replacement.Font.Bold = False 
 .Execute Forward:=True, Replace:=wdReplaceAll, _ 
 FindText:="", ReplaceWith:="" 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

