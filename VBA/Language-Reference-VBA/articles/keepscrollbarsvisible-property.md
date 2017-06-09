---
title: KeepScrollBarsVisible Property
keywords: fm20.chm5225048
f1_keywords:
- fm20.chm5225048
ms.prod: office
api_name:
- Office.KeepScrollBarsVisible
ms.assetid: e138df9f-5a8f-efcb-48db-9c8d22a7951f
ms.date: 06/08/2017
---


# KeepScrollBarsVisible Property



Specifies whether scroll bars remain visible when not required.
 **Syntax**
 _object_. **KeepScrollBarsVisible** [= _fmScrollBars_ ]
The  **KeepScrollBarsVisible** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmScrollBars_|Optional. Where scroll bars are displayed.|
 **Settings**
The settings for  _fmScrollBars_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmScrollBarsNone_|0|Displays no scroll bars.|
| _fmScrollBarsHorizontal_|1|Displays a horizontal scroll bar.|
| _fmScrollBarsVertical_|2|Displays a vertical scroll bar.|
| _fmScrollBarsBoth_|3|Displays both a horizontal and a vertical scroll bar (default).|
 **Remarks**
If the visible region is large enough to display all the controls on an object such as a  **Page** object or a form, scroll bars are not required. The **KeepScrollBarsVisible** property determines whether the scroll bars remain visible when they are not required.
If the scroll bars are visible when they are not required, they appear normal in size, and the scroll box fills the entire width or height of the scroll bar.

