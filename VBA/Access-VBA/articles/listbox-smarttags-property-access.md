---
title: ListBox.SmartTags Property (Access)
keywords: vbaac10.chm11303
f1_keywords:
- vbaac10.chm11303
ms.prod: access
api_name:
- Access.ListBox.SmartTags
ms.assetid: 1f35ca6b-fde1-6dc8-4b1b-f3089eee9204
ms.date: 06/08/2017
---


# ListBox.SmartTags Property (Access)

Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .


## Syntax

 _expression_. **SmartTags**

 _expression_ A variable that represents a **ListBox** object.


## Remarks




 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

