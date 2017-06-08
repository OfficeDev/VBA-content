---
title: Options.ParagraphAlignmentGuides Property (Word)
keywords: vbawd10.chm162988539
f1_keywords:
- vbawd10.chm162988539
ms.prod: word
ms.assetid: 2880558f-9ca3-73d6-cc2d-881efbab8c57
ms.date: 06/08/2017
---


# Options.ParagraphAlignmentGuides Property (Word)

Returns or sets a  **Boolean** that specifies whether paragraph alignment guides are displayed in the user interface. Read-write.


## Syntax

 _expression_ . **ParagraphAlignmentGuides**

 _expression_ A variable that represents an **Options** object.


## Remarks

If  **ParagraphAlignmentGuides** is set to **True**, paragraph alignment guides are displayed. Setting  **ParagraphAlignmentGuides** to **True** corresponds to selecting **Paragraph guides** under **Alignment Guides** in the **Grid and Guides** dialog box. (Click **Grid Settings** on the **Align** drop-down menu in the **Arrange** group on the **Format** contextual ribbon tab in the user interface.) For the **ParagraphAlignmentGuides** setting to have any effect, **[DisplayAlignmentGuides](options-displayalignmentguides-property-word.md)** must be set to **True**.


## Property value

 **BOOL**


## See also


#### Concepts


[Options Object](options-object-word.md)

