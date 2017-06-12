---
title: PrintRange.End Property (PowerPoint)
keywords: vbapp10.chm519004
f1_keywords:
- vbapp10.chm519004
ms.prod: powerpoint
api_name:
- PowerPoint.PrintRange.End
ms.assetid: 39f470c1-b469-3411-95e4-c6701487c498
ms.date: 06/08/2017
---


# PrintRange.End Property (PowerPoint)

Returns the number of the last slide in the specified print range. Read-only.


## Syntax

 _expression_. **End**

 _expression_ A variable that represents an **PrintRange** object.


### Return Value

Long


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

