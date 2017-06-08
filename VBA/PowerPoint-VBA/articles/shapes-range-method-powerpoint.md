---
title: Shapes.Range Method (PowerPoint)
keywords: vbapp10.chm543017
f1_keywords:
- vbapp10.chm543017
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Range
ms.assetid: 5ee926d9-5b30-a26b-7365-f4709a1a7bdb
ms.date: 06/08/2017
---


# Shapes.Range Method (PowerPoint)

Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents a subset of the shapes in a **[Shapes](shapes-object-powerpoint.md)** collection.


## Syntax

 _expression_. **Range**( **_Index_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The individual shapes that are to be included in the range. Can be an  **Integer** that specifies the index number of the shape, a **String** that specifies the name of the shape, or an array that contains either integers or strings. If this argument is omitted, the **Range** method returns all the objects in the specified collection.|

### Return Value

ShapeRange


## Remarks

Although you can use the  **Range** method to return any number of shapes or slides, it is simpler to use the **Item** method if you only want to return a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`, and  `Slides(2)` is simpler than `Slides.Range(2)`.

To specify an array of integers or strings for  **Index**, you can use the **Array** function. For example, the following instruction returns two shapes specified by name.

 `Dim myArray() As Variant, myRange As Object myArray = Array("Oval 4", "Rectangle 5") Set myRange = ActivePresentation.Slides(1).Shapes.Range(myArray)`


## Example

This example sets the fill pattern for shapes one and three on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Range(Array(1, 3)).Fill _
    .Patterned msoPatternHorizontalBrick
```

This example sets the fill pattern for the shapes named Oval 4 and Rectangle 5 on the first slide.




```vb
Dim myArray() As Variant, myRange As Object

myArray = Array("Oval 4", "Rectangle 5")

Set myRange = ActivePresentation.Slides(1).Shapes.Range(myArray)

myRange.Fill.Patterned msoPatternHorizontalBrick
```

This example sets the fill pattern for all shapes on the first slide.




```vb
ActivePresentation.Slides(1).Shapes.Range.Fill _
    .Patterned Pattern:=msoPatternHorizontalBrick
```

This example sets the fill pattern for shape one on the first slide.




```vb
Set myDocument = ActivePresentation.Slides(1)

Set myRange = myDocument.Shapes.Range(1)

myRange.Fill.Patterned msoPatternHorizontalBrick
```

This example creates an array that contains all the AutoShapes on the first slide, uses that array to define a shape range, and then distributes all the shapes in that range horizontally.




```vb
With myDocument.Shapes

    numShapes = .Count



    'Continues if there are shapes on the slide

    If numShapes > 1 Then

        numAutoShapes = 0

        ReDim autoShpArray(1 To numShapes)

        For i = 1 To numShapes



            'Counts the number of AutoShapes on the Slide

            If .Item(i).Type = msoAutoShape Then

                numAutoShapes = numAutoShapes + 1

                autoShpArray(numAutoShapes) = .Item(i).Name

            End If

        Next



        'Adds AutoShapes to ShapeRange

        If numAutoShapes > 1 Then

            ReDim Preserve autoShpArray(1 To numAutoShapes)

            Set asRange = .Range(autoShpArray)

            asRange.Distribute msoDistributeHorizontally, False

        End If

    End If

End With


```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

