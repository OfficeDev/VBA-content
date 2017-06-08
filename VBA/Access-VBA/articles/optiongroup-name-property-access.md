---
title: OptionGroup.Name Property (Access)
keywords: vbaac10.chm10871
f1_keywords:
- vbaac10.chm10871
ms.prod: access
api_name:
- Access.OptionGroup.Name
ms.assetid: 065ce5dd-8589-ee21-b850-d5ee95fb11ba
ms.date: 06/08/2017
---


# OptionGroup.Name Property (Access)

You can use the  **Name** property to specify or determine the string expression that identifies the name of an object. Read/write **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

A valid name must conform to the standard naming conventions for Microsoft Access. For Microsoft Access objects, the name may be up to 64 characters long. For controls, the name may be as long as 255 characters.

The default name for new objects is the object name plus a unique integer. For example, the first new form is Form1, the second new form is Form2, and so on. A form can't have the same name as another system object, such as the  **[Screen](screen-object-access.md)** object.

For an unbound control, the default name is the type of control plus a unique integer. For example, if the first control you add to a form is a text box, its  **Name** property setting is Text1

For a bound control, the default name is the name of the field in the underlying source of data. If you create a control by dragging a field from the field list, the field's  **FieldName** property setting is copied to the control's **Name** property box.

You can't use "Form" or "Report" to name a control or section.

Controls on the same form, report, or data access page can't have the same name, but controls on different forms, reports or data access pages can have the same name. A control and a section on the same form can't share the same name.


## See also


#### Concepts


[OptionGroup Object](optiongroup-object-access.md)

