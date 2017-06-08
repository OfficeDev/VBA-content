---
title: TextBox.ForeColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 6a6fb2a7-0434-d254-b825-4ce03c2a8db8
ms.date: 06/08/2017
---


# TextBox.ForeColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the foreground color of an object. Read/write.


## Syntax

 _expression_. **ForeColor**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```

Use the  **ForeColor** property for controls on forms to make them easy to read or to convey a special meaning. For example, if a text box reports the number of units in stock, you can change the color of the text when the value falls below the reorder level.


