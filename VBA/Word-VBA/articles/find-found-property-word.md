---
title: Find.Found Property (Word)
keywords: vbawd10.chm162529292
f1_keywords:
- vbawd10.chm162529292
ms.prod: word
api_name:
- Word.Find.Found
ms.assetid: c9a5d7ef-9df8-1439-248a-696c29fb01da
ms.date: 06/08/2017
---


# Find.Found Property (Word)

 **True** if the search produces a match. Read-only **Boolean** .


## Syntax

 _expression_ . **Found**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Example

This example removes formatting from the find criteria before searching the selection. If the word "Hello" with bold formatting is found, the entire paragraph is selected and copied to the Clipboard.


```vb
With Selection.Find 
 .ClearFormatting 
 .Font.Bold = True 
 .Execute FindText:="Hello", Format:=True, Forward:=True 
 If .Found = True Then 
 .Parent.Expand Unit:=wdParagraph 
 .Parent.Copy 
 End If 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

