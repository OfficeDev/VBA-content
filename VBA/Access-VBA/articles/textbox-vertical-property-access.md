---
title: TextBox.Vertical Property (Access)
keywords: vbaac10.chm11058
f1_keywords:
- vbaac10.chm11058
ms.prod: access
api_name:
- Access.TextBox.Vertical
ms.assetid: 40b9f9c0-daab-5562-395e-3e785d316d91
ms.date: 06/08/2017
---


# TextBox.Vertical Property (Access)

You can use the  **Vertical** property to set a form control for vertical display and editing or set a report control for vertical display and printing. Read/write **Boolean**.


## Syntax

 _expression_. **Vertical**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

The  **Vertical** property uses the following settings:



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|Displays, edits, and prints vertical text.|
|No|**False**|(Default) Does not display, edit, or print vertical text.|
You can specify how vertical text will be displayed, edited, or printed in the control by setting the  **Vertical** property. If set to Yes, the starting point for inputting text is the upper right corner of the control (the ending point is the lower left corner of the control). If using full pitch characters, the display and print directions are the same as the control for horizontal text. If using half pitch characters, it shifts 90 degrees to the right. The cursor is also rotated 90 degrees to the right in a vertical text control.


 **Note**  Text selection using key combinations is different for vertical text control and horizontal text control. Key combinations and their effect on range selection are described below.



|**Key combination**|**Selected range**|
|:-----|:-----|
|Shift+Up|Vertical: One character before the cursor. Horizontal: One line before the cursor.|
|Shift+Down|Vertical: One character after the cursor. Horizontal: One line after the cursor.|
|Shift+Right|Vertical: One line after the cursor. Horizontal: One character before the cursor.|
|Shift+Left|Vertical: One line before the cursor. Horizontal: One character after the cursor.|

## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

