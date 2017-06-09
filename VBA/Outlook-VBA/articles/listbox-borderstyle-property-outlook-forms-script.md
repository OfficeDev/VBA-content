---
title: ListBox.BorderStyle Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8b5996d0-4e03-f6e5-4418-0a28f4ac390d
ms.date: 06/08/2017
---


# ListBox.BorderStyle Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the type of border of the control. Read/write.


## Syntax

 _expression_. **BorderStyle**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The possible values of  **BorderStyle** are 0 and 1. 0 represents no visible border line, 1 represents a single-line border (default).

The default value for a  **[ListBox](listbox-object-outlook-forms-script.md)** is 0 (None).

You can use either  **BorderStyle** or **[SpecialEffect](listbox-specialeffect-property-outlook-forms-script.md)** to specify the border for a control, but not both. If you specify a nonzero value for one of these properties, the system sets the value of the other property to zero. For example, if you set **BorderStyle** to 1, the system sets **SpecialEffect** to zero (Flat). If you specify a nonzero value for **SpecialEffect**, the system sets  **BorderStyle** to zero.

 **BorderStyle** uses **[BorderColor](listbox-bordercolor-property-outlook-forms-script.md)** to define the colors of its borders. To use the **BorderColor** property, the **BorderStyle** property must be set to a value other than 0.


