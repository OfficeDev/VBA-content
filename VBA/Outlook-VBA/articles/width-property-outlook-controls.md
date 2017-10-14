---
title: Width Property (Outlook Controls)
keywords: olfm10.chm2002220
f1_keywords:
- olfm10.chm2002220
ms.prod: outlook
ms.assetid: 4457662c-930b-3849-d924-1607b35d8338
ms.date: 06/08/2017
---


# Width Property (Outlook Controls)

Returns or sets a  **Single** that specifies the width, in points, of the control. Read/write.


## Syntax

 _expression_. **Width**

 _expression_A variable that represents an Outlook control object.


## Remarks

The  **[Height](height-property-outlook-controls.md)** and **Width** properties are automatically updated when you move or size a control. If you specify a setting for the **[Left](left-property-outlook-controls.md)** or **[Top](top-property-outlook-controls.md)** property that is less than zero, that value will be used to calculate the height or width of the control, but a portion of the control will not be visible on the form.

If you move a control from one part of a form to another, the setting of  **Height** or **Width** only changes if you size the control as you move it. The settings of the control's **Left** and **Top** properties will change to reflect the control's new position relative to the edges of the form that contains it.

The value assigned to  **Height** or **Width** must be greater than or equal to zero. For most systems, the recommended range of values is from 0 to +32,767. Higher values may also work depending on your system configuration.

For most systems, the recommended range of values for  **Left** and **Top** is from -32,767 to +32,767. Other values may also work depending on your system configuration. For an **[OlkComboBox](olkcombobox-object-outlook.md)** control, values of **Left** and **Top** apply to the text box portion of the control, not to the list portion. When you move or size a control, its new **Left** setting is automatically entered in the property sheet.


