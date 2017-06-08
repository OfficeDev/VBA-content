---
title: SelText Property
keywords: fm20.chm2001890
f1_keywords:
- fm20.chm2001890
ms.prod: office
api_name:
- Office.SelText
ms.assetid: 75b9c27f-f6f7-6445-6d86-a53f046c1db6
ms.date: 06/08/2017
---


# SelText Property



Returns or sets the selected text of a control.
 **Syntax**
 _object_. **SelText** [= _String_ ]
The  **SelText** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression containing the selected text.|
 **Remarks**
If no characters are selected in the edit region of the control, the  **SelText** property returns a zero length string. This property is valid regardless of whether the control has the[focus](vbe-glossary.md).

