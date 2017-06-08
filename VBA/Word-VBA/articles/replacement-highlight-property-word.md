---
title: Replacement.Highlight Property (Word)
keywords: vbawd10.chm162594833
f1_keywords:
- vbawd10.chm162594833
ms.prod: word
api_name:
- Word.Replacement.Highlight
ms.assetid: 4bcccceb-7e0b-673d-09b7-d30a1938601e
ms.date: 06/08/2017
---


# Replacement.Highlight Property (Word)

 **True** if highlight formatting is applied to the replacement text. Read/write **Long** .


## Syntax

 _expression_ . **Highlight**

 _expression_ A variable that represents a **[Replacement](replacement-object-word.md)** object.


## Remarks

Can return or be set to  **True** , **False** , or **wdUndefined** .


## Example

This example finds all instances of highlighted text in the active document and removes the highlight formatting by setting the  **Highlight** property of the **Replacement** object to **False** .


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
With rngTemp.Find 
 .ClearFormatting 
 .Highlight = True 
 With .Replacement 
 .ClearFormatting 
 .Highlight = False 
 End With 
 .Execute Replace:=wdReplaceAll, Forward:=True, FindText:="", _ 
 ReplaceWith:="", Format:=True 
End With
```


## See also


#### Concepts


[Replacement Object](replacement-object-word.md)

