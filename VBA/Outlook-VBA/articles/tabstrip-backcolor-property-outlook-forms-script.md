---
title: TabStrip.BackColor Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 5d750acd-38aa-fa3d-55b3-3d10fb0d7c63
ms.date: 06/08/2017
---


# TabStrip.BackColor Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the background color of the object. Read/write.


## Syntax

 _expression_. **BackColor**

 _expression_A variable that represents a  **TabStrip** object.


## Remarks

You can use any integer that represents a valid color. You can also specify a color by using the Visual Basic  **RGB** function with red, green, and blue color components. The value of each color component is an integer that ranges from zero to 255. For example, you can specify teal blue as the integer value 4966415 or as red, green, and blue color components 15, 200, 75, as shown in the following example.


```
RGB(15,200,75)
```


