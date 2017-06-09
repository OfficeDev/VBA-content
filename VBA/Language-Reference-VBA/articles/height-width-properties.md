---
title: Height, Width Properties
keywords: fm20.chm916664
f1_keywords:
- fm20.chm916664
ms.prod: office
ms.assetid: b8fc82f0-c08f-c04a-58b7-062d8767e147
ms.date: 06/08/2017
---


# Height, Width Properties



The height or width, in [points](vbe-glossary.md), of an object.
 **Syntax**
 _object_. **Height** [= _Single_ ]
 _object_. **Width** [= _Single_ ]
The  **Height** and **Width** property syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Single_|Optional. A numeric expression specifying the dimensions of an object.|
 **Remarks**
The  **Height** and **Width** properties are automatically updated when you move or size a control. If you change the size of a control, the **Height** or **Width** property stores the new height or width and the **OldHeight** or **OldWidth** property stores the previous height or width. If you specify a setting for the **Left** or **Top** property that is less than zero, that value will be used to calculate the height or width of the control, but a portion of the control will not be visible on the form.
If you move a control from one part of a form to another, the setting of  **Height** or **Width** only changes if you size the control as you move it. The settings of the control's **Left** and **Top** properties will change to reflect the control's new position relative to the edges of the form that contains it.
The value assigned to  **Height** or **Width** must be greater than or equal to zero. For most systems, the recommended range of values is from 0 to +32,767. Higher values may also work depending on your system configuration.

