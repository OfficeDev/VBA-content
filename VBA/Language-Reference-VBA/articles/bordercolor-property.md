---
title: BorderColor Property
keywords: fm20.chm5225009
f1_keywords:
- fm20.chm5225009
ms.prod: office
api_name:
- Office.BorderColor
ms.assetid: f5718e93-55fa-e3c6-5359-c9ccc7c7a76c
ms.date: 06/08/2017
---


# BorderColor Property



Specifies the color of an object's border.
 **Syntax**
 _object_. **BorderColor** [= _Long_ ]
The  **BorderColor** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A value or constant that determines the border color of an object.|
 **Settings**
You can use any integer that represents a valid color. You can also specify a color by using the [RGB](glossary-vba.md) function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as RGB color component values 15, 200, 75.
 **Remarks**
To use the  **BorderColor** property, the **BorderStyle** property must be set to a value other than **fmBorderStyleNone**.
 **BorderStyle** uses **BorderColor** to define the border colors. The **SpecialEffect** property uses[system colors](glossary-vba.md) exclusively to define its border colors. For Windows operating systems, system color settings are part of the **Control Panel** and are found in the **Appearance** tab of the **Display** folder. In Windows NT 4.0 or later, system color settings are stored in the **Color** folder of the **Control Panel**.

