---
title: ParagraphFormat.SpaceAfterAuto Property (Word)
keywords: vbawd10.chm156434565
f1_keywords:
- vbawd10.chm156434565
ms.prod: word
api_name:
- Word.ParagraphFormat.SpaceAfterAuto
ms.assetid: c54c024a-5c04-fca5-95cb-bcbadb4baf41
ms.date: 06/08/2017
---


# ParagraphFormat.SpaceAfterAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceAfterAuto**

 _expression_ A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

Returns  **wdUndefined** if the **SpaceAfterAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceAfterAuto** is set to **True** , the **SpaceAfter** property is ignored.


## Example

This example displays a report showing the  **SpaceAfterAuto** settings for the active document.


```vb
Select Case ActiveDocument.Range.ParagraphFormat.SpaceAfterAuto 
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


[ParagraphFormat Object](paragraphformat-object-word.md)

