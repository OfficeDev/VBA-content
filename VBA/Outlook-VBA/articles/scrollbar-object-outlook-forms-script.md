---
title: ScrollBar Object (Outlook Forms Script)
keywords: olfm10.chm2000610
f1_keywords:
- olfm10.chm2000610
ms.prod: outlook
ms.assetid: 9e0a0f3d-fb04-2180-3beb-306b09c10c01
ms.date: 06/08/2017
---


# ScrollBar Object (Outlook Forms Script)

Returns or sets the value of another control based on the position of the scroll box.


## Remarks

A  **ScrollBar** is a stand-alone control you can place on a form. It is visually like the scroll bar you see in certain objects such as a **[ListBox](listbox-object-outlook-forms-script.md)** or the drop-down portion of a **[ComboBox](combobox-object-outlook-forms-script.md)**. However, unlike the scroll bars in these controls, the stand-alone  **ScrollBar** is not an integral part of any other control.

To use the  **ScrollBar** to set or read the value of another control, you must write code that uses the **ScrollBar** control's **[Value](scrollbar-value-property-outlook-forms-script.md)** property. For example, to use the **ScrollBar** to update the value of a **[TextBox](textbox-object-outlook-forms-script.md)**, you can write code that reads the  **Value** property of the **ScrollBar** and then sets the **[Value](scrollbar-value-property-outlook-forms-script.md)** property of the **TextBox**.

The default property for a  **ScrollBar** is the **Value** property.

To create a horizontal or vertical  **ScrollBar**, drag the sizing handles of the  **ScrollBar** horizontally or vertically on the form.


