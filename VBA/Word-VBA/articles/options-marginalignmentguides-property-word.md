---
title: Options.MarginAlignmentGuides Property (Word)
keywords: vbawd10.chm162988538
f1_keywords:
- vbawd10.chm162988538
ms.prod: word
ms.assetid: 0d5eed0b-4347-ff46-ed71-e2a025cae6ad
ms.date: 06/08/2017
---


# Options.MarginAlignmentGuides Property (Word)

Returns or sets a  **Boolean** that specifies whether margin alignment guides are displayed in the user interface. Read-write.


## Syntax

 _expression_ . **MarginAlignmentGuides**

 _expression_ A variable that represents an **Options** object.


## Remarks

If  **MarginAlignmentGuides** is set to **True**, margin alignment guides are displayed. Setting  **MarginAlignmentGuides** to **True** corresponds to selecting **Margin guides** under **Alignment Guides** in the **Grid and Guides** dialog box. (Click **Grid Settings** on the **Align** drop-down menu in the **Arrange** group on the **Format** contextual ribbon tab in the user interface.) For the **MarginAlignmentGuides** setting to have any effect, **[DisplayAlignmentGuides](options-displayalignmentguides-property-word.md)** must be set to **True**.


## Property value

 **BOOL**


## See also


#### Concepts


[Options Object](options-object-word.md)

