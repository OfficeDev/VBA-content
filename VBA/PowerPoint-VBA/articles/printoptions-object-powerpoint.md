---
title: PrintOptions Object (PowerPoint)
keywords: vbapp10.chm517000
f1_keywords:
- vbapp10.chm517000
ms.prod: powerpoint
api_name:
- PowerPoint.PrintOptions
ms.assetid: 19ce56ba-b0d0-4086-db86-e32feade70bd
ms.date: 06/08/2017
---


# PrintOptions Object (PowerPoint)

Contains print options for a presentation.


 **Note**  Specifying the optional arguments From, To, Copies, and Collate for the  **[PrintOut](presentation-printout-method-powerpoint.md)** method sets the corresponding properties of the **PrintOptions** object.


## Example

Use the [PrintOptions](presentation-printoptions-property-powerpoint.md) property to return the **PrintOptions** object. The following example prints two uncollated color copies of all the slides (whether visible or hidden) in the active presentation. The example also scales each slide to fit the printed page and frames each slide with a thin border.


```vb
With ActivePresentation 
    With .PrintOptions 
        .NumberOfCopies = 2 
        .Collate = False 
        .PrintColorType = ppPrintColor 
        .PrintHiddenSlides = True 
        .FitToPage = True 
        .FrameSlides = True 
        .OutputType = ppPrintOutputSlides 
    End With 
    .PrintOut 
End With
```

Use the [RangeType](printoptions-rangetype-property-powerpoint.md) property to specify whether to print the entire presentation or only a specified part of it. If you want to print only certain slides, set the **RangeType** property to **ppPrintSlideRange**, and use the[Ranges](printoptions-ranges-property-powerpoint.md) property to specify which pages to print. The following example prints slides 1, 4, 5, and 6 in the active presentation




```vb
With ActivePresentation 
    With .PrintOptions 
        .RangeType = ppPrintSlideRange 
        With .Ranges 
            .Add 1, 1 
            .Add 4, 6 
        End With 
    End With 
    .PrintOut 
End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

