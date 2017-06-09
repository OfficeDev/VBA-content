---
title: QBColor Function
keywords: vblr6.chm1012946
f1_keywords:
- vblr6.chm1012946
ms.prod: office
ms.assetid: fa9c0598-8454-fd05-a34c-be5e25158816
ms.date: 06/08/2017
---


# QBColor Function



Returns a [Long](vbe-glossary.md) representing the RGB color code corresponding to the specified color number.
 **Syntax**
 **QBColor(**_color_**)**
The required  _color_[argument](vbe-glossary.md) is a whole number in the range 0-15.
 **Settings**
The  _color_ argument has these settings:


|**Number**|**Color**|**Number**|**Color**|
|:-----|:-----|:-----|:-----|
|0|Black|8|Gray|
|1|Blue|9|Light Blue|
|2|Green|10|Light Green|
|3|Cyan|11|Light Cyan|
|4|Red|12|Light Red|
|5|Magenta|13|Light Magenta|
|6|Yellow|14|Light Yellow|
|7|White|15|Bright White|
 **Remarks**
The  _color_ argument represents color values used by earlier versions of Basic (such as Microsoft Visual Basic for MS-DOS and the Basic Compiler). Starting with the least-significant byte, the returned value specifies the red, green, and blue values used to set the appropriate color in the RGB system used by Visual Basic for Applications.

## Example

This example uses the  **QBColor** function to change the **BackColor** property of the form passed in as `MyForm` to the color indicated by `ColorCode`.  **QBColor** accepts integer values between 0 and 15.


```vb
Sub ChangeBackColor (ColorCode As Integer, MyForm As Form)
    MyForm.BackColor = QBColor(ColorCode)
End Sub
```


