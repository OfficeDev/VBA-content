---
title: CommandButton.Enabled Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 51520f48-79c6-2bfe-7c21-ba41baf46f22
ms.date: 06/08/2017
---


# CommandButton.Enabled Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.


## Syntax

 _expression_. **Enabled**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

 **True** is the control can receive the focus and respond to user-generated events, and is accessible through code (default). **False** if the user cannot interact with the control by using the mouse, keystrokes, accelerators, or hotkeys. The control is generally still accessible through code.

Use the  **Enabled** property to enable and disable controls. A disabled control appears dimmed, while an enabled control does not. Also, if a control displays a bitmap, the bitmap is dimmed whenever the control is dimmed.

The  **Enabled** and **[Locked](commandbutton-locked-property-outlook-forms-script.md)** properties work together to achieve the following effects:


- If  **Enabled** and **Locked** are both **True**, the control can receive focus and appears normally (not dimmed) in the form. The user can copy, but not edit, data in the control.
    
- If  **Enabled** is **True** and **Locked** is **False**, the control can receive focus and appears normally in the form. The user can copy and edit data in the control.
    
- If  **Enabled** is **False** and **Locked** is **True**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    
- If  **Enabled** and **Locked** are both **False**, the control cannot receive focus and is dimmed in the form. The user can neither copy nor edit data in the control.
    


You can combine the settings of the  **Enabled** and the **TabStop** properties to prevent the user from selecting a **[CommandButton](commandbutton-object-outlook-forms-script.md)** with TAB, while still allowing the user to click the button. Setting **TabStop** to **False** means that the **CommandButton** won't appear in the tab order. However, if **Enabled** is **True**, then the user can still click the command button, as long as  **[TakeFocusOnClick](commandbutton-takefocusonclick-property-outlook-forms-script.md)** is set to **True**.


