---
title: Paragraph.SpaceBeforeAuto Property (Word)
keywords: vbawd10.chm156696708
f1_keywords:
- vbawd10.chm156696708
ms.prod: word
api_name:
- Word.Paragraph.SpaceBeforeAuto
ms.assetid: 4c69088a-fcc2-ee0f-dfb5-74491d0b1737
ms.date: 06/08/2017
---


# Paragraph.SpaceBeforeAuto Property (Word)

 **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .


## Syntax

 _expression_ . **SpaceBeforeAuto**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

This property returns  **wdUndefined** if the **SpaceBeforeAuto** property is set to **True** for only some of the specified paragraphs. Can be set to either **True** or **False** .

If  **SpaceBeforeAuto** is set to **True** , the **SpaceBefore** property is ignored.


## Example

This example displays a report showing the  **SpaceBeforeAuto** settings for the first paragraph in the active document.


```vb
Select Case ActiveDocument.Paragraphs(1).SpaceBeforeAuto 
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


[Paragraph Object](paragraph-object-word.md)

