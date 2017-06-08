---
title: ScrollBar.LargeChange Property (Outlook Forms Script)
keywords: olfm10.chm2001360
f1_keywords:
- olfm10.chm2001360
ms.prod: outlook
ms.assetid: 1236ef08-7788-a345-e2a6-a3c647fe2675
ms.date: 06/08/2017
---


# ScrollBar.LargeChange Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the amount of movement that occurs when the user clicks between the scroll box and scroll arrow. Read/write.


## Syntax

 _expression_. **LargeChange**

 _expression_A variable that represents a  **ScrollBar** object.


## Remarks

The  **LargeChange** property specifies the amount of change to the **[Value](scrollbar-value-property-outlook-forms-script.md)** property.

The  **LargeChange** property applies only to the **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. It does not apply to the scrollbars in other controls such as a  **[TextBox](textbox-object-outlook-forms-script.md)** or a drop-down **[ComboBox](combobox-object-outlook-forms-script.md)**.

The value of  **LargeChange** is the amount by which the **ScrollBar** control's **Value** property changes when the user clicks the area between the scroll box and scroll arrow. The direction of the movement is always toward the place where the user clicks. For example, in a horizontal **ScrollBar**, clicking to the left of the scroll box moves the scroll box to the left. In a vertical  **ScrollBar**, clicking above the scroll box moves the scroll box up.

 **LargeChange** does not have units. Any integer is a valid setting for **LargeChange**. The recommended range of values is from -32,767 to +32,767, and the value must be between the values of the  **[Max](scrollbar-max-property-outlook-forms-script.md)** and **[Min](scrollbar-min-property-outlook-forms-script.md)** properties of the **ScrollBar**.


