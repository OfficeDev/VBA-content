---
title: SetFocus Method
keywords: fm20.chm5224972
f1_keywords:
- fm20.chm5224972
ms.prod: office
api_name:
- Office.SetFocus
ms.assetid: 430b2404-f11f-a0b6-e3b7-4bfe513c9258
ms.date: 06/08/2017
---


# SetFocus Method



Moves the [focus](vbe-glossary.md) to this instance of an object.
 **Syntax**
 _object_. **SetFocus**
The  **SetFocus** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
If setting the focus fails, the focus reverts to the previous object and an error is generated.
By default, setting the focus to a control does not activate the control's window or place it on top of other controls.
The  **SetFocus** method is valid for an empty **Frame** as well as a **Frame** that contains other controls. An empty **Frame** will take the focus itself, and any subsequent keyboard events apply to the **Frame**. In a **Frame** that contains other controls, the focus moves to the first control in the **Frame**, and subsequent keyboard events apply to the control that has the focus.

