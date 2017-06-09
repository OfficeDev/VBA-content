---
title: ComboBox.BorderStyle Property (Outlook Forms Script)
keywords: olfm10.chm2000820
f1_keywords:
- olfm10.chm2000820
ms.prod: outlook
ms.assetid: 59caf8ee-9287-362e-1102-c40a9f61bf8d
ms.date: 06/08/2017
---


# ComboBox.BorderStyle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the type of border of the control. Read/write.


## Syntax

 _expression_. **BorderStyle**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

The possible values of  **BorderStyle** are 0 and 1. 0 represents no visible border line, 1 represents a single-line border (default).

The default value for a  **[ComboBox](combobox-object-outlook-forms-script.md)** is 0 (None).

You can use either  **BorderStyle** or **[SpecialEffect](combobox-specialeffect-property-outlook-forms-script.md)** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to 1, the system sets **SpecialEffect** to zero (Flat). If you specify a nonzero value for **SpecialEffect**, the system sets  **BorderStyle** to zero.

 **BorderStyle** uses **[BorderColor](combobox-bordercolor-property-outlook-forms-script.md)** to define the colors of its borders. To use the **BorderColor** property, the **BorderStyle** property must be set to a value other than 0.


