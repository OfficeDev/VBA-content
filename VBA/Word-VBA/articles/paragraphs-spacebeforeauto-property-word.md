---
title: Paragraphs.SpaceBeforeAuto Property (Word)
keywords: vbawd10.chm156762244
f1_keywords:
- vbawd10.chm156762244
ms.prod: word
api_name:
- Word.Paragraphs.SpaceBeforeAuto
ms.assetid: be2bbab2-81bb-a95e-201b-46487fda2ca8
ms.date: 06/08/2017
---


# Paragraphs.SpaceBeforeAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceBeforeAuto**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if the **SpaceBeforeAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceBeforeAuto** is set to **True** , the **SpaceBefore** property is ignored.


## Example

This example displays a report showing the  **SpaceBeforeAuto** settings for the active document.


```vb
Select Case ActiveDocument.Paragraphs.SpaceBeforeAuto 
 Case True 
 x = "Spacing before paragraphs is handled " _ 
 &; "automatically for all paragraphs." 
 Case False 
 x = "Spacing before paragraphs is handled " _ 
 &; "manually for all paragraphs." 
 Case wdUndefined 
 x = "Spacing before paragraphs is handled " _ 
 &; "automatically for some paragraphs, " _ 
 &; "manually for some paragraphs." 
End Select
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

