---
title: ParagraphFormat.SpaceBeforeAuto Property (Word)
keywords: vbawd10.chm156434564
f1_keywords:
- vbawd10.chm156434564
ms.prod: word
api_name:
- Word.ParagraphFormat.SpaceBeforeAuto
ms.assetid: c3c86ee1-c62f-d921-2dc7-d7201b181622
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceBeforeAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceBeforeAuto**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

Returns  **wdUndefined** if the **SpaceBeforeAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceBeforeAuto** is set to **True** , the **SpaceBefore** property is ignored.


## Example

This example displays a report showing the  **SpaceBeforeAuto** settings for the active document.


```vb
Select Case ActiveDocument.Range.ParagraphFormat.SpaceBeforeAuto 
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


[ParagraphFormat Object](paragraphformat-object-word.md)

