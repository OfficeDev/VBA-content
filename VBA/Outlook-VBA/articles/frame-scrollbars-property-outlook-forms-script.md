---
title: Frame.ScrollBars Property (Outlook Forms Script)
keywords: olfm10.chm2001780
f1_keywords:
- olfm10.chm2001780
ms.prod: outlook
ms.assetid: 2a443602-40f7-6f56-0626-479fcd0efd38
ms.date: 06/08/2017
---


# Frame.ScrollBars Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether a control has vertical scroll bars, horizontal scroll bars, or both. Read/write.


## Syntax

 _expression_. **ScrollBars**

 _expression_A variable that represents a  **Frame** object.


## Remarks

The settings for  **ScrollBars** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Displays no scroll bars (default).|
|1|Displays a horizontal scroll bar.|
|2|Displays a vertical scroll bar.|
|3|Displays both a horizontal and a vertical scroll bar.|
If the  **[KeepScrollBarsVisible](frame-keepscrollbarsvisible-property-outlook-forms-script.md)** property is **True**, any scroll bar on a form or page is always visible, regardless of whether the object's contents fit within the object's borders.

If visible, a scroll bar constrains its scroll box to the visible region of the scroll bar. It also modifies the scroll position as needed to keep the entire scroll bar visible. The range of a scroll bar changes when the value of the  **ScrollBars** property changes, the scroll size changes, or the visible size changes.

If a scroll bar is not visible, then you can set its scroll position to any value. Negative values and values greater than the scroll size are both valid.


