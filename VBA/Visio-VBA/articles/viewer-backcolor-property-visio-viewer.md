---
title: Viewer.BackColor Property (Visio Viewer)
ms.prod: visio
ms.assetid: 8718d3b6-9521-85d3-64fc-32feeb118274
ms.date: 06/08/2017
---


# Viewer.BackColor Property (Visio Viewer)

Gets or sets the background color of Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **[OLE_COLOR]**


## Remarks

Returns a value of data type  **OLE_COLOR** that represents the background color of Visio Viewer. The **OLE_COLOR** data type is used for properties that return colors.

Valid hexadecimal values for an  **OLE_COLOR** data type in Visio Viewer are of the form _&;Hbbggrr_, where  _bb_ is the blue value, _gg_ the green value, and _rr_ the red value. All three color values range between 00 and FF hexadecimal (255 decimal).

The  **BackColor** property controls the color shown in the Visio Viewer window behind the images shown for page and shapes. The default value of the **BackColor** property matches the color of the current Windows color scheme, if that color is available; otherwise, the default is white. To set **BackColor** to "Visio blue," use the hexadecimal value &;HFFFFA0 (or the decimal value 16777120).


## Example

The following code sets the value of the  **BackColor** property to the default value in a Windows form.


```
 vsoViewer.BackColor = &;H8000000F
```


