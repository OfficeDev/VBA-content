---
title: NavigationButton.ObjectPalette Property (Access)
keywords: vbaac10.chm10490
f1_keywords:
- vbaac10.chm10490
ms.prod: access
api_name:
- Access.NavigationButton.ObjectPalette
ms.assetid: 10578730-717c-6c3c-d6d4-61a9bc765ca3
ms.date: 06/08/2017
---


# NavigationButton.ObjectPalette Property (Access)

The  **ObjectPalette** property specifies the palette in the application used to create A bitmap or other graphic that is loaded into the specified control by using the **Picture** property. Read/write **Variant**.


## Syntax

 _expression_. **ObjectPalette**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

Microsoft Access sets the value of the  **ObjectPalette** property to a **String** data type containing the palette information. You can use this setting to set the value of the **PaintPalette** property for a form or report.

If the application associated with the bitmap or other graphic doesn't have an associated palette, the  **ObjectPalette** property is set to an zero-length string.

The  **ObjectPalette** property is read-only in Form Design view, Form view, and Report Design view. This property setting is unavailable in other views.

The setting of the  **ObjectPalette** property makes the palette of the application associated with the OLE object contained in a control available to the **PaintPalette** property of a form or report. For example, to make the palette used in Microsoft Graph available when you're designing a form in Microsoft Access, you set the form's **PaintPalette** property to the **ObjectPalette** value of an existing chart control.


 **Note**   Windows can have only one color palette active at a time. Microsoft Access allows you to have multiple graphics on a form, each using a different color palette. The **PaintPalette** and **PaletteSource** properties let you specify which color palette a form should use when displaying graphics.


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

