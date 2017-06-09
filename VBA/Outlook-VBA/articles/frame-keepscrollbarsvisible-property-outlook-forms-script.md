---
title: Frame.KeepScrollBarsVisible Property (Outlook Forms Script)
keywords: olfm10.chm2001350
f1_keywords:
- olfm10.chm2001350
ms.prod: outlook
ms.assetid: adc2bda2-6e7f-cd02-c6ca-f2976250fd60
ms.date: 06/08/2017
---


# Frame.KeepScrollBarsVisible Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether scroll bars remain visible when not required. Read/write.


## Syntax

 _expression_. **KeepScrollBarsVisible**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The settings for  **KeepScrollBarsVisible** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Displays no scroll bars.|
|1|Displays a horizontal scroll bar.|
|2|Displays a vertical scroll bar.|
|3|Displays both a horizontal and a vertical scroll bar (default).|
If the visible region is large enough to display all the controls on an object such as a  **[Page](page-object-outlook-forms-script.md)** object, scroll bars are not required. The **KeepScrollBarsVisible** property determines whether the scroll bars remain visible when they are not required.

If the scroll bars are visible when they are not required, they appear normal in size, and the scroll box fills the entire width or height of the scroll bar.

If the  **KeepScrollBarsVisible** property is **True**, any scroll bar on a form or page is always visible, regardless of whether the object's contents fit within the object's borders.


