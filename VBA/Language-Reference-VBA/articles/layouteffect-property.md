---
title: LayoutEffect Property
keywords: fm20.chm5225050
f1_keywords:
- fm20.chm5225050
ms.prod: office
api_name:
- Office.LayoutEffect
ms.assetid: 74e19b13-605c-caa8-4a12-e877d638d316
ms.date: 06/08/2017
---


# LayoutEffect Property



Specifies whether a control was moved during a layout change.
 **Syntax**
 _object_. **LayoutEffect**
The  **LayoutEffect** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Return Values**
The  **LayoutEffect** property return values are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmLayoutEffectNone_|0|The control was not moved.|
| _fmLayoutEffectInitiate_|1|The control moved.|
 **Remarks**
The  **LayoutEffect** property is read-only and is available only in the Layout event. The Layout event is initiated by the **Move** method if the _Layout_ argument is **True**.
The Layout event is not initiated when you change the settings of the  **Left**, **Top**, **Height**, or **Width** properties of a control.
The Layout event sets  **LayoutEffect** for any control that was involved in a move operation. For example, if you move a group of controls, **LayoutEffect** of each control is set.

