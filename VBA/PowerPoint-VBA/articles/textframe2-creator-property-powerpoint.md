---
title: TextFrame2.Creator Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.Creator
ms.assetid: e591a997-2322-cf14-d79b-0b63aa9d9e46
ms.date: 06/08/2017
---


# TextFrame2.Creator Property (PowerPoint)

Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. Read-only.


## Syntax

 _expression_. **Creator**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

Long


## Remarks

For example, if a  **TextFrame2** object was created in PowerPoint, this property returns the hexadecimal number 50575054.

The  **Creator** property is designed to be used in Microsoft Office applications for the Macintosh.


## Example

This example displays a message about the creator of the  **TextFrame2** object.


```vb
Public Sub Creator_Example()



    Set pptTextFrame2 = Application.ActivePresentation.Slides(1).Shapes(1).TextFrame2

    If pptTextFrame2.Creator = &;H50575054 Then

        MsgBox "This is a PowerPoint object"

    Else

        MsgBox "This is not a PowerPoint object"

    End If

    

End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

