---
title: Report.PaletteSource Property (Access)
keywords: vbaac10.chm13759
f1_keywords:
- vbaac10.chm13759
ms.prod: access
api_name:
- Access.Report.PaletteSource
ms.assetid: 9dc324a1-dc31-b0c5-edca-c4bc1674155a
ms.date: 06/08/2017
---


# Report.PaletteSource Property (Access)

You can use the  **PaletteSource** property to specify the palette for a report. Read/write **String**.


## Syntax

 _expression_. **PaletteSource**

 _expression_ A variable that represents a **Report** object.


## Remarks

Enter the path and file name of one of the following file types:


- .dib (device-independent bitmap file)
    
- .pal (Windows palette file)
    
- .ico (Windows icon file)
    
- .bmp (Windows bitmap file)
    
- .wmf or .emf file, or other graphics file for which you have a graphics filter
    
The default setting is (Default), which specifies the palette included with Microsoft Access. If you change this setting by entering a path and file name, the property setting displays (Custom).

For a report, you can set the  **PaletteSource** property only in report Design view. The property setting is unavailable in other views.

Windows can have only one color palette active at a time. Microsoft Access allows you to have multiple graphics on a form, each using a different color palette. The  **PaletteSource** and **PaintPalette** properties let you specify which color palette a form uses when displaying graphics.

When you set the  **PaletteSource** property for a form or report, Microsoft Access automatically updates its **PaintPalette** property. Conversely, when you set a form's or report's **PaintPalette** property, the **PaletteSource** property is also updated. For example, when you specify a custom palette with the **PaintPalette** property, the **PaletteSource** property setting changes to (Custom). The **PaintPalette** property (which is available only in a macro or Visual Basic) is used to set the palette for the form or report. The **PaletteSource** property gives you a way to set the palette for the form or report in the property sheet by using an existing graphics file.


## Example

The following example sets the  **PaintPalette** property of the Seascape form to the **ObjectPalette** property of the Ocean control on the DisplayPictures form. (Ocean can be a bound object frame, command button, chart, toggle button, or unbound object frame.)


```vb
Forms!Seascape.PaintPalette = _ 
 Forms!DisplayPictures!Ocean.ObjectPalette
```

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


[Report Object](report-object-access.md)

