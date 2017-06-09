---
title: AutoSize Property
keywords: fm20.chm5225003
f1_keywords:
- fm20.chm5225003
ms.prod: office
api_name:
- Office.AutoSize
ms.assetid: 581fd371-1ab4-78fd-1f04-e775bc576125
ms.date: 06/08/2017
---


# AutoSize Property



Specifies whether an object automatically resizes to display its entire contents.
 **Syntax**
 _object_. **AutoSize** [= _Boolean_ ]
The  **AutoSize** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control is resized.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|Automatically resizes the control to display its entire contents.|
|**False**|Keeps the size of the control constant. Contents are clipped when they exceed the area of the control (default).|
 **Remarks**
For controls with captions, the  **AutoSize** property specifies whether the control automatically adjusts to display the entire caption.
For controls without captions, this property specifies whether the control automatically adjusts to display the information stored in the control. In a  **ComboBox**, for example, setting **AutoSize** to **True** automatically sets the width of the display area to match the length of the current text.
For a single-line text box, setting  **AutoSize** to **True** automatically sets the width of the display area to the length of the text in the text box.
For a multiline text box that contains no text, setting  **AutoSize** to **True** automatically displays the text as a column. The width of the text column is set to accommodate the widest letter of that font size. The height of the text column is set to display the entire text of the **TextBox**.
For a multiline text box that contains text, setting  **AutoSize** to **True** automatically enlarges the **TextBox** vertically to display ithe entire text. The width of the **TextBox** does not change.

 **Note**  If you manually change the size of a control while  **AutoSize** is **True**, the manual change overrides the size previously set by **AutoSize**.


