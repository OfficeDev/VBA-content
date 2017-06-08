---
title: Page.ScrollBars Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2a4c0132-9d91-c1cb-3e95-061e12012c81
ms.date: 06/08/2017
---


# Page.ScrollBars Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether a page has vertical scroll bars, horizontal scroll bars, or both. Read/write.


## Syntax

 _expression_. **ScrollBars**

 _expression_A variable that represents a  **Page** object.


## Remarks

The settings for  **ScrollBars** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Displays no scroll bars (default).|
|1|Displays a horizontal scroll bar.|
|2|Displays a vertical scroll bar.|
|3|Displays both a horizontal and a vertical scroll bar.|
If the  **[KeepScrollBarsVisible](page-keepscrollbarsvisible-property-outlook-forms-script.md)** property is **True**, any scroll bar on a form or page is always visible, regardless of whether the object's contents fit within the object's borders.

If visible, a scroll bar constrains its scroll box to the visible region of the scroll bar. It also modifies the scroll position as needed to keep the entire scroll bar visible. The range of a scroll bar changes when the value of the  **ScrollBars** property changes, the scroll size changes, or the visible size changes.

If a scroll bar is not visible, then you can set its scroll position to any value. Negative values and values greater than the scroll size are both valid.


