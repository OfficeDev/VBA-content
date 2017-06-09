---
title: Paragraphs.SpaceAfterAuto Property (Word)
keywords: vbawd10.chm156762245
f1_keywords:
- vbawd10.chm156762245
ms.prod: word
api_name:
- Word.Paragraphs.SpaceAfterAuto
ms.assetid: 699b6a20-63dd-55f1-a0da-f26a3a1f7bfc
ms.date: 06/08/2017
---


# Paragraphs.SpaceAfterAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceAfterAuto**

 _expression_ A variable that represents a **[Paragraphs](paragraphs-object-word.md)** collection.


## Remarks

This property returns  **wdUndefined** if the **SpaceAfterAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceAfterAuto** is set to **True** , the **SpaceAfter** property is ignored.


## Example

This example displays a report showing the  **SpaceAfterAuto** settings for the active document.


```vb
Select Case ActiveDocument.Paragraphs.SpaceAfterAuto 
 Case True 
 x = "Spacing after paragraphs is handled " _ 
 &; "automatically for all paragraphs." 
 Case False 
 x = "Spacing after paragraphs is handled " _ 
 &; "manually for all paragraphs." 
 Case wdUndefined 
 x = "Spacing after paragraphs is handled " _ 
 &; "automatically for some paragraphs, " _ 
 &; "manually for some paragraphs." 
End Select
```


## See also


#### Concepts


[Paragraphs Collection Object](paragraphs-object-word.md)

