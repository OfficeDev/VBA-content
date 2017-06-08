---
title: ChartArea.ClearFormats Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartArea.ClearFormats
ms.assetid: 80732262-f84d-1153-811e-30ce887a8661
ms.date: 06/08/2017
---


# ChartArea.ClearFormats Method (PowerPoint)

Clears the formatting of the object.


## Syntax

 _expression_. **ClearFormats**

 _expression_ A variable that represents a **[ChartArea](chartarea-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example clears the formatting from the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartArea.ClearFormats

    End If

End With
```


## See also


#### Concepts


[ChartArea Object](chartarea-object-powerpoint.md)

