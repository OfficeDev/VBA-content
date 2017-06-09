---
title: OptionButton.GroupName Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 7b3b92d6-e1e0-6171-4d6a-4b0221e1c083
ms.date: 06/08/2017
---


# OptionButton.GroupName Property (Outlook Forms Script)

Returns or sets a  **String** that identifies a group of mutually exclusive **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls. Read/write.


## Syntax

 _expression_. **GroupName**

 _expression_A variable that represents an  **OptionButton** object.


## Remarks

Use the same  **GroupName** for all buttons in the group. The default setting is an empty string.

To create a group of mutually exclusive  **OptionButton** controls, you can put the buttons in a **[Frame](frame-object-outlook-forms-script.md)** on your form, or you can use the **GroupName** property. **GroupName** is more efficient for the following reasons:


- You do not have to include a  **Frame** for each group. By not using a **Frame**, you reduce the number of controls on the form, and in turn, improve performance and reduce the size of the form.
    
- You have more design flexibility. If you use a  **Frame** to create the group, all the buttons must be inside the **Frame**. If you want more than one group, you must have one  **Frame** for each group. However, if you use **GroupName** to create the group, the group can include option buttons anywhere on the form. If you want more than one group, specify a unique name for each group; you can still place the individual controls anywhere on the form.
    
- You can create buttons with transparent backgrounds, which can improve the visual appearance of your form. The  **Frame** is not a transparent control.
    


Regardless of which method you use to create the group of buttons, clicking one button in a group sets all other buttons in the same group to  **False**. All option buttons with the same  **GroupName** within a single container are mutually exclusive. You can use the same group name in two containers, but doing so creates two groups (one in each container) rather than one group that includes both containers.

For example, assume your form includes some option buttons and a  **[MultiPage](multipage-object-outlook-forms-script.md)** that also includes option buttons. The option buttons on the **MultiPage** are one group and the buttons on the form are another group. The two groups do not affect each other. Changing the setting of a button on the **MultiPage** does not affect the buttons on the form.


