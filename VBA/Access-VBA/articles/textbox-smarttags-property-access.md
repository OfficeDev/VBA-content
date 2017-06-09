---
title: TextBox.SmartTags Property (Access)
keywords: vbaac10.chm11148
f1_keywords:
- vbaac10.chm11148
ms.prod: access
api_name:
- Access.TextBox.SmartTags
ms.assetid: 200175d1-78a2-3036-72ba-4a85dfc21864
ms.date: 06/08/2017
---


# TextBox.SmartTags Property (Access)

Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .


## Syntax

 _expression_. **SmartTags**

 _expression_ A variable that represents a **TextBox** object.


## Remarks




 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

