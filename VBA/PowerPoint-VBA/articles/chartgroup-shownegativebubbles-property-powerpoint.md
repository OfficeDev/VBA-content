---
title: ChartGroup.ShowNegativeBubbles Property (PowerPoint)
keywords: vbapp10.chm692002
f1_keywords:
- vbapp10.chm692002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.ShowNegativeBubbles
ms.assetid: 0f197a05-0f3c-841f-e3f7-a27c2e8b6130
ms.date: 06/08/2017
---


# ChartGroup.ShowNegativeBubbles Property (PowerPoint)

 **True** if negative bubbles are shown for the chart group. Read/write **Boolean**.


## Syntax

 _expression_. **ShowNegativeBubbles**

 _expression_ A variable that represents a **[ChartGroup](chartgroup-object-powerpoint.md)** object.


## Remarks

The property is valid only for bubble charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes Microsoft Word to display negative bubbles for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).ShowNegativeBubbles = True

    End If

End With
```


## See also


#### Concepts


[ChartGroup Object](chartgroup-object-powerpoint.md)

