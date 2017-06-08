---
title: CheckBox.GroupName Property (Outlook Forms Script)
keywords: olfm10.chm2001245
f1_keywords:
- olfm10.chm2001245
ms.prod: outlook
ms.assetid: 24fc2e67-273d-2ef3-1040-e5fa9161bd81
ms.date: 06/08/2017
---


# CheckBox.GroupName Property (Outlook Forms Script)

Returns or sets a  **String** that identifies a group of mutually exclusive **[CheckBox](checkbox-object-outlook-forms-script.md)** controls. Read/write.


## Syntax

 _expression_. **GroupName**

 _expression_A variable that represents a  **CheckBox** object.


## Remarks

Use the same  **GroupName** for all check boxes in the group. The default setting is an empty string.

To create a group of mutually exclusive  **CheckBox** controls, you can put the check boxes in a **[Frame](frame-object-outlook-forms-script.md)** on your form, or you can use the **GroupName** property. **GroupName** is more efficient for the following reasons:


- You do not have to include a  **Frame** for each group. By not using a **Frame**, you reduce the number of controls on the form, and in turn, improve performance and reduce the size of the form.
    
- You have more design flexibility. If you use a  **Frame** to create the group, all the check boxes must be inside the **Frame**. If you want more than one group, you must have one  **Frame** for each group. However, if you use **GroupName** to create the group, the group can include check boxes anywhere on the form. If you want more than one group, specify a unique name for each group; you can still place the individual controls anywhere on the form.
    
- You can create check boxes with transparent backgrounds, which can improve the visual appearance of your form. The  **Frame** is not a transparent control.
    


Regardless of which method you use to create the group of check boxes, clicking one check box in a group sets all other check boxes in the same group to  **False**. All check boxes with the same  **GroupName** within a single container are mutually exclusive. You can use the same group name in two containers, but doing so creates two groups (one in each container) rather than one group that includes both containers.

For example, assume your form includes some check boxes and a  **[MultiPage](multipage-object-outlook-forms-script.md)** that also includes option buttons. The check boxes on the **MultiPage** are one group and the buttons on the form are another group. The two groups do not affect each other. Changing the setting of a check box on the **MultiPage** does not affect the check boxes on the form.


