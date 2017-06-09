---
title: Find.Replacement Property (Word)
keywords: vbawd10.chm162529305
f1_keywords:
- vbawd10.chm162529305
ms.prod: word
api_name:
- Word.Find.Replacement
ms.assetid: b0c728d6-4f2e-6c01-da95-ab59c79ce752
ms.date: 06/08/2017
---


# Find.Replacement Property (Word)

Returns a  **[Replacement](replacement-object-word.md)** object that contains the criteria for a replace operation.


## Syntax

 _expression_ . **Replacement**

 _expression_ An expression that returns a **[Find](find-object-word.md)** object.


## Example

This example removes bold formatting from the active document. The  **Bold** property of the **Font** object is **True** for the **Find** object and **False** for the **Replacement** object.


```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 End With 
 .Execute FindText:="", ReplaceWith:="", Format:=True, _ 
 Replace:=wdReplaceAll 
End With
```

This example finds every instance of the word "Start" in the active document and replaces it with "End." The find operation ignores formatting but matches the case of the text to find ("Start").




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
With myRange.Find 
 .ClearFormatting 
 .Text = "Start" 
 With .Replacement 
 .ClearFormatting 
 .Text = "End" 
 End With 
 .Execute Replace:=wdReplaceAll, _ 
 Format:=True, MatchCase:=True, _ 
 MatchWholeWord:=True 
End With
```


## See also


#### Concepts


[Find Object](find-object-word.md)

