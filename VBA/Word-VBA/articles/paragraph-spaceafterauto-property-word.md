---
title: Paragraph.SpaceAfterAuto Property (Word)
keywords: vbawd10.chm156696709
f1_keywords:
- vbawd10.chm156696709
ms.prod: word
api_name:
- Word.Paragraph.SpaceAfterAuto
ms.assetid: ca17c146-ad99-2d2c-8f04-4c6183bf7182
ms.date: 06/08/2017
---


# Paragraph.SpaceAfterAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceAfterAuto**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

Returns  **wdUndefined** if the **SpaceAfterAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceAfterAuto** is set to **True** , the **SpaceAfter** property is ignored.


## Example

This example displays a report showing the  **SpaceAfterAuto** settings for the first paragraph in the active document.


```vb
Select Case ActiveDocument.Paragraphs(1).SpaceAfterAuto 
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


[Paragraph Object](paragraph-object-word.md)

