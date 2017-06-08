---
title: Control.Name Property (Access)
keywords: vbaac10.chm10173
f1_keywords:
- vbaac10.chm10173
ms.prod: access
api_name:
- Access.Control.Name
ms.assetid: b1e31997-1b99-0476-eda8-afef8975420b
ms.date: 06/08/2017
---


# Control.Name Property (Access)

You can use the Name property to determine the string expression that identifies the name of an object. Read-only  **String**.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Control** object.


## Remarks

A valid name must conform to the standard naming conventions for Microsoft Access. For controls, the name may be as long as 255 characters.

For an unbound control, the default name is the type of control plus a unique integer. For example, if the first control you add to a form is a text box, its  **Name** property setting is Text1

For a bound control, the default name is the name of the field in the underlying source of data. If you create a control by dragging a field from the field list, the field's  **FieldName** property setting is copied to the control's **Name** property box.

You can't use "Form" or "Report" to name a control or section.

Controls on the same form, report, or data access page can't have the same name, but controls on different forms, reports or data access pages can have the same name. A control and a section on the same form can't share the same name.


## See also


#### Concepts


[Control Object](control-object-access.md)

