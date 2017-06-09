---
title: RGB Function
keywords: vblr6.chm1009005
f1_keywords:
- vblr6.chm1009005
ms.prod: office
ms.assetid: 5e9956de-ba18-56cd-0556-715774055cf4
ms.date: 06/08/2017
---


# RGB Function



Returns a [Long](vbe-glossary.md) whole number representing an RGB color value.
 **Syntax**
 **RGB( _red_, _green_, _blue_ )**
The  **RGB** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_red_**|Required;  **Variant** ( **Integer** ). Number in the range 0-255, inclusive, that represents the red component of the color.|
|**_green_**|Required;  **Variant** ( **Integer** ). Number in the range 0-255, inclusive, that represents the green component of the color.|
|**_blue_**|Required;  **Variant** ( **Integer** ). Number in the range 0-255, inclusive, that represents the blue component of the color.|
 **Remarks**
Application [methods](vbe-glossary.md) and[properties](vbe-glossary.md) that accept a color specification expect that specification to be a number representing an RGB color value. An RGB color value specifies the relative intensity of red, green, and blue to cause a specific color to be displayed.
The value for any [argument](vbe-glossary.md) to **RGB** that exceeds 255 is assumed to be 255.
The following table lists some standard colors and the red, green, and blue values they include:


|**Color**|**Red Value**|**Green Value**|**Blue Value**|
|:-----|:-----|:-----|:-----|
|Black|0|0|0|
|Blue|0|0|255|
|Green|0|255|0|
|Cyan|0|255|255|
|Red|255|0|0|
|Magenta|255|0|255|
|Yellow|255|255|0|
|White|255|255|255|
The RGB color values returned by this function are incompatible with those used by the Macintosh operating system. They may be used within the context of Microsoft applications for the Macintosh, but should not be used when communicating color changes directly to the Macintosh operating system.

## Example

This example shows how the  **RGB** function is used to return a whole number representing an **RGB** color value. It is used for those application methods and properties that accept a color specification. The object `MyObject` and its property are used for illustration purposes only. If `MyObject` does not exist, or if it does not have a **Color** property, an error occurs.


```vb
Dim RED, I, RGBValue, MyObject
Red = RGB(255, 0, 0)    ' Return the value for Red.
I = 75    ' Initialize offset.
RGBValue = RGB(I, 64 + I, 128 + I)     ' Same as RGB(75, 139, 203).
MyObject.Color = RGB(255, 0, 0)    ' Set the Color property of 
    ' MyObject to Red.

```


