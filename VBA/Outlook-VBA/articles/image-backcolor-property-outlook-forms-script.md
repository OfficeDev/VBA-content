---
title: Image.BackColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: c0064240-d0f8-3bb3-fb93-c758a0f749e4
ms.date: 06/08/2017
---


# Image.BackColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the background color of the object. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_A variable that represents an  **Image** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```

You can only see the background color of an object if the  **[BackStyle](image-backstyle-property-outlook-forms-script.md)** property is set to 1.


