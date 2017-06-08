---
title: SpinButton.BackColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 64de2a16-04a8-2a27-96a9-51bcd5962e2d
ms.date: 06/08/2017
---


# SpinButton.BackColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the background color of the object. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_A variable that represents a  **SpinButton** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```


