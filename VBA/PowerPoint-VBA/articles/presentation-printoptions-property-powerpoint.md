---
title: Presentation.PrintOptions Property (PowerPoint)
keywords: vbapp10.chm583033
f1_keywords:
- vbapp10.chm583033
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.PrintOptions
ms.assetid: 3620e0bb-1dcc-9979-d815-c3f34205aaaf
ms.date: 06/08/2017
---


# Presentation.PrintOptions Property (PowerPoint)

Returns a  **[PrintOptions](printoptions-object-powerpoint.md)** object that represents print options that are saved with the specified presentation. Read-only.


## Syntax

 _expression_. **PrintOptions**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

PrintOptions


## Example

This example causes hidden slides in the active presentation to be printed, and it scales the printed slides to fit the paper size.


```vb
With Application.ActivePresentation

    With .PrintOptions

        .PrintHiddenSlides = True

        .FitToPage = True

    End With

    .PrintOut

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

