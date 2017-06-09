---
title: Frame.BorderColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 348a2dd5-0b16-327a-0a83-124b338d4b44
ms.date: 06/08/2017
---


# Frame.BorderColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the border color of an object. Read/write.


## Syntax

 _expression_. **BorderColor**

 _expression_A variable that represents a  **Frame** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```

To use the  **BorderColor** property, the **[BorderStyle](frame-borderstyle-property-outlook-forms-script.md)** property must be set to a value other than 0.

 **BorderStyle** uses **BorderColor** to define the border colors. The **[SpecialEffect](frame-specialeffect-property-outlook-forms-script.md)** property uses system colors exclusively to define its border colors. For Windows operating systems, system color settings are set using the **Display** icon in Control Panel.


