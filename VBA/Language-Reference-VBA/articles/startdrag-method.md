---
title: StartDrag Method
keywords: fm20.chm5224974
f1_keywords:
- fm20.chm5224974
ms.prod: office
api_name:
- Office.StartDrag
ms.assetid: 9713f582-759f-2cb2-825f-a79469041dc8
ms.date: 06/08/2017
---


# StartDrag Method



Initiates a drag-and-drop operation for a  **DataObject**.
 **Syntax**
 _fmDropEffect=Object_. **StartDrag _(_**_[Effect as fmDropEffect])_
The  **StartDrag** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _Object_|Required. A valid object.|
| _Effect_|Optional. Effect of the drop operation on the target control.|
 **Settings**
The settings for  _Effect_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmDropEffectNone_|0|Does not copy or move the [drop source](glossary-vba.md) to the drop target.|
| _fmDropEffectCopy_|1|Copies the drop source to the drop target.|
| _fmDropEffectMove_|2|Moves the drop source to the drop target.|
| _fmDropEffectCopyOrMove_|3|Copies or moves the drop source to the drop target.|
 **Remarks**
The drag action starts at the current mouse pointer position with the current [keyboard state](glossary-vba.md) and ends when the user releases the mouse. The effect of the drag-and-drop operation depends on the effect chosen for the drop target.
For example, a control's MouseMove event might include the  **StartDrag** method. When the user clicks the control and moves the mouse, the mouse pointer changes to indicate whether _Effect_ is valid for the drop target.

