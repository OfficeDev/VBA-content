---
title: ComboBox.SmartTags Property (Access)
keywords: vbaac10.chm11478
f1_keywords:
- vbaac10.chm11478
ms.prod: ACCESS
api_name:
- Access.ComboBox.SmartTags
ms.assetid: b86a8460-48c6-92ad-602b-1d736bb2c38c
---


# ComboBox.SmartTags Property (Access)

Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .


## Syntax

 _expression_. **SmartTags**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks




 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

