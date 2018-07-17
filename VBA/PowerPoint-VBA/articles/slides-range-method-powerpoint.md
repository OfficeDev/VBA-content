---
title: Slides.Range Method (PowerPoint)
keywords: vbapp10.chm530007
f1_keywords:
- vbapp10.chm530007
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.Range
ms.assetid: f3950ce5-7873-86e8-5625-7ad2a0cb77dd
ms.date: 06/08/2017
---


# Slides.Range Method (PowerPoint)

Returns a  **[SlideRange](sliderange-object-powerpoint.md)** object that represents a subset of the slides in a **[Slides](slides-object-powerpoint.md)** collection.


## Syntax

 _expression_. **Range**( **_Index_** )

 _expression_ A variable that represents a **Slides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The individual slides that are to be included in the range. Can be an  **Integer** that specifies the index number of the slide, a **String** that specifies the name of the slide, or an array that contains either integers or strings. If this argument is omitted, the **Range** method returns all the objects in the specified collection.|

### Return Value

SlideRange


## Remarks

Although you can use the  **Range** method to return any number of shapes or slides, it is simpler to use the **Item** method if you only want to return a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`, and  `Slides(2)` is simpler than `Slides.Range(2)`.

To specify an array of integers or strings for  **Index**, you can use the **Array** function. For example, the following instruction returns two shapes specified by name.

 `Dim myArray() As Variant, myRange As Object myArray = Array("Oval 4", "Rectangle 5") Set myRange = ActivePresentation.Slides(1).Shapes.Range(myArray)`


## Example

This example sets the title color for slides one and three.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```

This example sets the title color for the slides named Slide6 and Slide8.




```vb
Set mySlides = ActivePresentation.Slides _
    .Range(Array("Slide6", "Slide8"))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```

This example sets the title color for all the slides in the active presentation.




```vb
Set mySlides = ActivePresentation.Slides.Range

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(255, 0, 0)
```

This example creates an array that contains all the title slides in the active presentation, uses that array to define a slide range, and then sets the title color for all slides in that range.




```vb
Dim MyTitleArray() As Long

Set pSlides = ActivePresentation.Slides

ReDim MyTitleArray(1 To pSlides.Count)

For Each pSlide In pSlides

    If pSlide.Layout = ppLayoutTitle Then

        nCounter = nCounter + 1

        MyTitleArray(nCounter) = pSlide.SlideIndex

    End If

Next pSlide

ReDim Preserve MyTitleArray(1 To nCounter)



Set rngTitleSlides = ActivePresentation.Slides.Range(MyTitleArray)

rngTitleSlides.ColorScheme.Colors(ppTitle).RGB = RGB(255, 123, 99)
```


## See also


#### Concepts


[Slides Object](slides-object-powerpoint.md)

