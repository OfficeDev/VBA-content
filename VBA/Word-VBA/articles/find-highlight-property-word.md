---
title: Find.Highlight Property (Word)
keywords: vbawd10.chm162529304
f1_keywords:
- vbawd10.chm162529304
ms.prod: word
api_name:
- Word.Find.Highlight
ms.assetid: 75873be2-035e-ae93-1f5d-28e86d383a8c
ms.date: 06/08/2017
---


# Find.Highlight Property (Word)

 **True** if highlight formatting is included in the find criteria. Read/write **Long** .


## Syntax

 _expression_ . **Highlight**

 _expression_ A variable that represents a **[Find](find-object-word.md)** object.


## Remarks

The  **Highlight** property can return or be set to **True** , **False** , or **wdUndefined** . The **wdUndefined** value can be used with the **Find** object to ignore the state of highlight formatting in the selection or range that is searched.


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


[Find Object](find-object-word.md)

