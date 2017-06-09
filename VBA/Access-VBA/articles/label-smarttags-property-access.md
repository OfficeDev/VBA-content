---
title: Label.SmartTags Property (Access)
keywords: vbaac10.chm10242
f1_keywords:
- vbaac10.chm10242
ms.prod: access
api_name:
- Access.Label.SmartTags
ms.assetid: 1c31246b-870d-2d73-1737-829cbd67baba
ms.date: 06/08/2017
---


# Label.SmartTags Property (Access)

Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .


## Syntax

 _expression_. **SmartTags**

 _expression_ A variable that represents a **Label** object.


## Remarks




 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## See also


#### Concepts


[Label Object](label-object-access.md)

