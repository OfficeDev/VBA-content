---
title: TextBox.ScrollBars Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: ef258f24-df52-bdf5-6c1e-04b8b41d4c7f
ms.date: 06/08/2017
---


# TextBox.ScrollBars Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies whether a control has vertical scroll bars, horizontal scroll bars, or both. Read/write.


## Syntax

 _expression_. **ScrollBars**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

The settings for  **ScrollBars** are:



|**Value**|**Description**|
|:-----|:-----|
|0|Displays no scroll bars (default).|
|1|Displays a horizontal scroll bar.|
|2|Displays a vertical scroll bar.|
|3|Displays both a horizontal and a vertical scroll bar.|
A scroll bar constrains its scroll box to the visible region of the scroll bar. It also modifies the scroll position as needed to keep the entire scroll bar visible. The range of a scroll bar changes when the value of the  **ScrollBars** property changes, the scroll size changes, or the visible size changes.

For a single-line control, you can display a horizontal scroll bar by using the  **ScrollBars** and **[AutoSize](textbox-autosize-property-outlook-forms-script.md)** properties. Scroll bars are hidden or displayed according to the following rules:


1. When  **ScrollBars** is set to 0, no scroll bar is displayed.
    
2. When  **ScrollBars** is set to 1 or 3, the control displays a horizontal scroll bar if the text is longer than the edit region and if the control has enough room to include the scroll bar underneath its edit region.
    
3. When  **AutoSize** is **True**, the control enlarges itself to accommodate the addition of a scroll bar unless the control is at or near its maximum size.
    


For a multiline  **[TextBox](textbox-object-outlook-forms-script.md)**, you can display scroll bars by using the  **ScrollBars**,  **[WordWrap](textbox-wordwrap-property-outlook-forms-script.md)**, and  **AutoSize** properties. Scroll bars are hidden or displayed according to the following rules:


1. When  **ScrollBars** is set to 0, no scroll bar is displayed.
    
2. When  **ScrollBars** is set to 2 or 3, the control displays a vertical scroll bar if the text is longer than the edit region and if the control has enough room to include the scroll bar at the right edge of its edit region.
    
3. When  **WordWrap** is **True**, the multiline control will not display a horizontal scroll bar. Most multiline controls do not use a horizontal scroll bar.
    
4. A multiline control can display a horizontal scroll bar if the following conditions occur simultaneously:
    
      - The edit region contains a word that is longer than the edit region's width.
    
  - The control has enabled horizontal scroll bars.
    
  - The control has enough room to include the scroll bar under the edit region.
    
  - The  **WordWrap** property is set to **False**.
    



