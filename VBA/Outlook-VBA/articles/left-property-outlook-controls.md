---
title: Left Property (Outlook Controls)
keywords: olfm10.chm2001380
f1_keywords:
- olfm10.chm2001380
ms.prod: outlook
ms.assetid: 1fe354f7-5a4e-ba65-c12c-13498a22bdc2
ms.date: 06/08/2017
---


# Left Property (Outlook Controls)

Returns or sets a  **Single** that represents the distance between a control and the left edge of the form that contains it. Read/write.


## Syntax

 _expression_. **Left**

 _expression_A variable that represents an Outlook control object.


## Remarks

Setting the  **Left** or **[Top](top-property-outlook-controls.md)** property to 0 places the control's edge at the left or top edge of its container.

For most systems, the recommended range of values for  **Left** and **Top** is from -32,767 to +32,767. Other values may also work depending on your system configuration. For an **[OlkComboBox](olkcombobox-object-outlook.md)** control, values of **Left** and **Top** apply to the text box portion of the control, not to the list portion. When you move or size a control, its new **Left** setting is automatically entered in the property sheet.

The  **[Height](height-property-outlook-controls.md)** and **[Width](width-property-outlook-controls.md)** properties are automatically updated when you move or size a control. If you specify a setting for the **Left** or **Top** property that is less than zero, that value will be used to calculate the height or width of the control, but a portion of the control will not be visible on the form.

If you move a control from one part of a form to another, the setting of  **Height** or **Width** only changes if you size the control as you move it. The settings of the control's **Left** and **Top** properties will change to reflect the control's new position relative to the edges of the form that contains it.


