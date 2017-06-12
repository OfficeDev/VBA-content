---
title: Options.PageAlignmentGuides Property (Word)
keywords: vbawd10.chm162988537
f1_keywords:
- vbawd10.chm162988537
ms.prod: word
ms.assetid: 027ff389-7288-e5c8-b437-5e6c650ccdf6
ms.date: 06/08/2017
---


# Options.PageAlignmentGuides Property (Word)

Returns or sets a  **Boolean** that specifies whether page alignment guides are displayed in the user interface. Read-write.


## Syntax

 _expression_ . **PageAlignmentGuides**

 _expression_ A variable that represents an **Options** object.


## Remarks

If  **PageAlignmentGuides** is set to **True**, page alignment guides are displayed at the edges of the page. Setting  **PageAlignmentGuides** to **True** corresponds to selecting **Page guides** under **Alignment Guides** in the **Grid and Guides** dialog box. (Click **Grid Settings** on the **Align** drop-down menu in the **Arrange** group on the **Format** contextual ribbon tab in the user interface.) For the **PageAlignmentGuides** setting to have any effect, **[DisplayAlignmentGuides](options-displayalignmentguides-property-word.md)** must be set to **True**.


## Property value

 **BOOL**


## See also


#### Concepts


[Options Object](options-object-word.md)

