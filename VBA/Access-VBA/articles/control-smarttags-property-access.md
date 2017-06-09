---
title: Control.SmartTags Property (Access)
keywords: vbaac10.chm10153
f1_keywords:
- vbaac10.chm10153
ms.prod: access
api_name:
- Access.Control.SmartTags
ms.assetid: 2f8b1435-31d4-4388-614c-4f26544eed7c
ms.date: 06/08/2017
---


# Control.SmartTags Property (Access)

Returns a  **[SmartTags](smarttags-object-access.md)** collection that represents the collection of smart tags that have been added to a control. .


## Syntax

 _expression_. **SmartTags**

 _expression_ A variable that represents a **Control** object.


## Remarks

Using the  **SmartTags** property will result in a run-time error if the control is not a **ComboBox**, **Label**, or **TextBox** object.


 **Note**  Unlike the  **SmartTags** collections in Microsoft Excel and Microsoft Word, the **SmartTags** collection in Microsoft Access is zero-based. Therefore, the code `control.SmartTags(0)` returns the first smart tag for the specified control.


## See also


#### Concepts


[Control Object](control-object-access.md)

