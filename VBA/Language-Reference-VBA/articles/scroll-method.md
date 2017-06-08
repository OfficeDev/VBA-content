---
title: Scroll Method
keywords: fm20.chm2000390
f1_keywords:
- fm20.chm2000390
ms.prod: office
api_name:
- Office.Scroll
ms.assetid: dbbfcf37-c511-3112-55f6-b2e8ca055db3
ms.date: 06/08/2017
---


# Scroll Method



Moves the scroll bar on an object.
 **Syntax**
 _object_. **Scroll(** [ _ActionX_ [, _ActionY_ ]] **)**
The  **Scroll** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _ActionX_|Optional. Identifies the action to occur in the horizontal direction.|
| _ActionY_|Optional. Identifies the action to occur in the vertical direction.|
 **Settings**
The settings for  _ActionX_ and _ActionY_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmScrollActionNoChange_|0|Do not scroll in the specified direction.|
| _fmScrollActionLineUp_|1|Move up on a vertical scroll bar or left on a horizontal scroll bar. Movement is equivalent to pressing the up or left arrow key on the keyboard to move the scroll bar.|
| _fmScrollActionLineDown_|2|Move down on a vertical scroll bar or right on a horizontal scroll bar. Movement is equivalent to pressing the right or down arrow key on the keyboard to move the scroll bar.|
| _fmScrollActionPageUp_|3|Move one pageup on a vertical scroll bar or one page left on a horizontal scroll bar. Movement is equivalent to pressing PAGE UP on the keyboard to move the scroll bar.|
| _fmScrollActionPageDown_|4|Move one pagedown on a vertical scroll bar or one page right on a horizontal scroll bar. Movement is equivalent to pressing PAGE DOWN on the keyboard to move the scroll bar.|
| _fmScrollActionBegin_|5|Move to the top of a vertical scroll bar or to the left end of a horizontal scroll bar.|
| _fmScrollActionEnd_|6|Move to the bottom of a vertical scroll bar or to the right end of a horizontal scroll bar.|
 **Remarks**
The  **Scroll** method applies scroll bars that appear on a form, **Frame**, or **Page** that is larger than its display area. This method does not apply to the stand-alone **ScrollBar** or to scroll bars that appear on a **TextBox**.

