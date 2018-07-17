---
title: Form.PaintPalette Property (Access)
keywords: vbaac10.chm13426
f1_keywords:
- vbaac10.chm13426
ms.prod: access
api_name:
- Access.Form.PaintPalette
ms.assetid: 161a7bfa-c861-68b9-eaac-05a2d7c24d4a
ms.date: 06/08/2017
---


# Form.PaintPalette Property (Access)

You can use the  **PaintPalette** property to specify a palette to be used by a form. Read/write **Variant**.


## Syntax

 _expression_. **PaintPalette**

 _expression_ A variable that represents a **Form** object.


## Remarks

You can set the  **PaintPalette** property by using a macro or Visual Basic . The property setting must be a **String** data type containing the palette information.

You can set the  **PaintPalette** property by assigning the value of the **ObjectPalette** property to the **PaintPalette** property in a macro or Visual Basic, by setting the **PaletteSource** property (in which case Microsoft Access automatically sets the **PaintPalette** property to this **PaletteSource** ), or by setting the **PaintPalette** property of one form or report to the **PaintPalette** property of another form or report.

For a form, you can set the  **PaintPalette** property in form Design view and Form view.

When you set the  **PaintPalette** property, Microsoft Access makes a copy of the palette that you specify and saves it with the form or report. The palette is then available if you modify that form or report.

Changes to the palette you specified when you set the  **PaintPalette** property don't affect the copy of the palette stored with the form or report. If you want to update the copy of the palette stored with the form or report, you must rerun the code or macro that sets the **PaintPalette** property or reset the **PaletteSource** property when the form or report is open.

When you set the  **PaintPalette** property for a form or report, Microsoft Access automatically updates its **PaletteSource** property. Conversely, when you set the **PaletteSource** property for a form or report, the **PaintPalette** property is also updated. For example, when you specify a custom palette with the **PaintPalette** property, the **PaletteSource** property setting is changed to (Custom). The **PaintPalette** property (which is available only in a macro or Visual Basic) is used to set the palette for the form or report. The **PaletteSource** property gives you a way to set the palette for the form or report in the property sheet by using an existing graphics file.


 **Note**  Windows can have only one color palette active at a time. Microsoft Access allows you to have multiple graphics on a form, each using a different color palette. The  **PaintPalette** and **PaletteSource** properties let you specify which color palette a form should use when displaying graphics.

You can use the  **ObjectPalette** property to make the palette of an application associated with an OLE object, bitmap, or other graphic contained in a control on a form or report available to the **PaintPalette** property. For example, to make the palette used in Microsoft Graph available when you're designing a form in Microsoft Access, you set the form's **PaintPalette** property to the **ObjectPalette** value of an existing chart control.


## Example

The  **ObjectPalette** and **PaintPalette** properties are useful for programmatically altering the color palette in use by an open form at run time. A common use of these properties is to set the current form's **PaintPalette** property to the palette of a graphic displayed in a control that has the focus.

For example, you can have a form with an ocean picture, showing many shades of blue, and a sunset picture, showing many shades of red. Since Windows only allows one color palette active at a time, one picture will look much better than the other. The following example uses a control's Enter event for setting the form's  **PaintPalette** property to the control's **ObjectPalette** property so the graphic that has the focus will have an optimal appearance.




```vb
Sub OceanPicture_Enter() 
 Me.PaintPalette = Me!OceanPicture.ObjectPalette 
End Sub 
 
Sub SunsetPicture_Enter() 
 Me.PaintPalette = Me!SunsetPicture.ObjectPalette 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

