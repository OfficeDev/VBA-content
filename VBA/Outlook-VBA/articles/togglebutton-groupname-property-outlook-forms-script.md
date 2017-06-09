---
title: ToggleButton.GroupName Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 69787bc9-90cb-c2f7-380d-2f48ab2db270
ms.date: 06/08/2017
---


# ToggleButton.GroupName Property (Outlook Forms Script)

Returns or sets a  **String** that identifies a group of mutually exclusive **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls. Read/write.


## Syntax

 _expression_. **GroupName**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

Use the same  **GroupName** for all buttons in the group. The default setting is an empty string.

To create a group of mutually exclusive  **ToggleButton** controls, you can put the buttons in a **[Frame](frame-object-outlook-forms-script.md)** on your form, or you can use the **GroupName** property. **GroupName** is more efficient for the following reasons:


- You do not have to include a  **Frame** for each group. By not using a **Frame**, you reduce the number of controls on the form, and in turn, improve performance and reduce the size of the form.
    
- You have more design flexibility. If you use a  **Frame** to create the group, all the buttons must be inside the **Frame**. If you want more than one group, you must have one  **Frame** for each group. However, if you use **GroupName** to create the group, the group can include toggle buttons anywhere on the form. If you want more than one group, specify a unique name for each group; you can still place the individual controls anywhere on the form.
    
- You can create buttons with transparent backgrounds, which can improve the visual appearance of your form. The  **Frame** is not a transparent control.
    


Regardless of which method you use to create the group of buttons, clicking one button in a group sets all other buttons in the same group to  **False**. All toggle buttons with the same  **GroupName** within a single container are mutually exclusive. You can use the same group name in two containers, but doing so creates two groups (one in each container) rather than one group that includes both containers.

For example, assume your form includes some toggle buttons and a  **[MultiPage](multipage-object-outlook-forms-script.md)** that also includes toggle buttons. The toggle buttons on the **MultiPage** are one group and the buttons on the form are another group. The two groups do not affect each other. Changing the setting of a button on the **MultiPage** does not affect the buttons on the form.


