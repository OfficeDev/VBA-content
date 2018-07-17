---
title: HideSelection Property
keywords: fm20.chm2001270
f1_keywords:
- fm20.chm2001270
ms.prod: office
api_name:
- Office.HideSelection
ms.assetid: fe840b76-7f50-8801-642f-3cce6707bb62
ms.date: 06/08/2017
---


# HideSelection Property



Specifies whether selected text remains highlighted when a control does not have the [focus](vbe-glossary.md).
 **Syntax**
 _object_. **HideSelection** [= _Boolean_ ]
The  **HideSelection** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the selected text remains highlighted even when the control does not have the focus.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Selected text is not highlighted unless the control has the focus (default).|
|**False**|Selected text always appears highlighted.|
 **Remarks**
You can use the  **HideSelection** property to maintain highlighted text when another form or a dialog box receives the focus, such as in a spell-checking procedure.

