---
title: ScrollBars Property
keywords: fm20.chm2001780
f1_keywords:
- fm20.chm2001780
ms.prod: office
api_name:
- Office.ScrollBars
ms.assetid: cf925c0d-45ac-a724-6149-2aed7725b593
ms.date: 06/08/2017
---


# ScrollBars Property



Specifies whether a control, form, or page has vertical scroll bars, horizontal scroll bars, or both.
 **Syntax**
 _object_. **ScrollBars** [= _fmScrollBars_ ]
The  **ScrollBars** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmScrollBars_|Optional. Where scroll bars should be displayed.|
 **Settings**
The settings for  _fmScrollBars_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmScrollBarsNone_|0|Displays no scroll bars (default).|
| _fmScrollBarsHorizontal_|1|Displays a horizontal scroll bar.|
| _fmScrollBarsVertical_|2|Displays a vertical scroll bar.|
| _fmScrollBarsBoth_|3|Displays both a horizontal and a vertical scroll bar.|
 **Remarks**
If the  **KeepScrollBarsVisible** property is **True**, any scroll bar on a form or page is always visible, regardless of whether the object's contents fit within the object's borders.
If visible, a scroll bar constrains its scroll box to the visible region of the scroll bar. It also modifies the scroll position as needed to keep the entire scroll bar visible. The range of a scroll bar changes when the value of the  **ScrollBars** property changes, the scroll size changes, or the visible size changes.
If a scroll bar is not visible, then you can set its scroll position to any value. Negative values and values greater than the scroll size are both valid.
For a single-line control, you can display a horizontal scroll bar by using the  **ScrollBars** and **AutoSize** properties. Scroll bars are hidden or displayed according to the following rules:


1. When  **ScrollBars** is set to **fmScrollBarsNone**, no scroll bar is displayed.
    
2. When  **ScrollBars** is set to **fmScrollBarsHorizontal** or **fmScrollBarsBoth**, the control displays a horizontal scroll bar if the text is longer than the edit region and if the control has enough room to include the scroll bar underneath its edit region.
    
3. When  **AutoSize** is **True**, the control enlarges itself to accommodate the addition of a scroll bar unless the control is at or near its maximum size.
    

For a multiline  **TextBox**, you can display scroll bars by using the **ScrollBars**, **WordWrap**, and **AutoSize** properties. Scroll bars are hidden or displayed according to the following rules:


1. When  **ScrollBars** is set to **fmScrollBarsNone**, no scroll bar is displayed.
    
2. When  **ScrollBars** is set to **fmScrollBarsVertical** or **fmScrollBarsBoth**, the control displays a vertical scroll bar if the text is longer than the edit region and if the control has enough room to include the scroll bar at the right edge of its edit region.
    
3. When  **WordWrap** is **True**, the multiline control will not display a horizontal scroll bar. Most multiline controls do not use a horizontal scroll bar.
    
4. A multiline control can display a horizontal scroll bar if the following conditions occur simultaneously:
    
    
    
      - The edit region contains a word that is longer than the edit region's width.
    
  - The control has enabled horizontal scroll bars.
    
  - The control has enough room to include the scroll bar under the edit region.
    
  - The  **WordWrap** property is set to **False**.
    

    
    


