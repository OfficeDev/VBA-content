---
title: PrintRange.Start Property (PowerPoint)
keywords: vbapp10.chm519003
f1_keywords:
- vbapp10.chm519003
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRange.Start
ms.assetid: 493d64b3-c2fb-7f4a-ca59-a7f657a386a0
ms.date: 06/08/2017
---


# PrintRange.Start Property (PowerPoint)

Returns the number of the first slide in the range of slides to be printed. Read-only.


## Syntax

 _expression_. **Start**

 _expression_ A variable that represents a **PrintRange** object.


### Return Value

Integer


## Example

This example displays a message that indicates the starting and ending slide numbers for print range one in the active presentation.


```vb
With ActivePresentation.PrintOptions.Ranges
    If .Count > 0 Then
        With .Item(1)
            MsgBox "Print range 1 starts on slide " &; .Start &; _
                " and ends on slide " &; .End
        End With
    End If
End With
```


## See also


#### Concepts


[PrintRange Object](printrange-object-powerpoint.md)

