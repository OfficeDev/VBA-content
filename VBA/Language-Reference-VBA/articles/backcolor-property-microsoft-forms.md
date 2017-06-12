---
title: BackColor Property (Microsoft Forms)
keywords: fm20.chm2000770
f1_keywords:
- fm20.chm2000770
ms.prod: office
ms.assetid: 70549eaf-d785-67e7-3f04-76151864d850
ms.date: 06/08/2017
---


# BackColor Property (Microsoft Forms)



Specifies the [background color](glossary-vba.md) of the object.
 **Syntax**
 _object_. **BackColor** [= _Long_ ]
The  **BackColor** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. A value or constant that determines the background color of an object.|
 **Settings**
You can use any integer that represents a valid color. You can also specify a color by using the [RGB](glossary-vba.md) function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75.
 **Remarks**
You can only see the background color of an object if the  **BackStyle** property is set to **fmBackStyleOpaque**.

