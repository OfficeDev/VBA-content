---
title: ShapeRange.HasTextFrame Property (PowerPoint)
keywords: vbapp10.chm548055
f1_keywords:
- vbapp10.chm548055
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.HasTextFrame
ms.assetid: 6359fbeb-0a91-ad56-9edd-9b6be7fe51b7
ms.date: 06/08/2017
---


# ShapeRange.HasTextFrame Property (PowerPoint)

Returns whether the specified shape has a text frame. Read-only.


## Syntax

 _expression_. **HasTextFrame**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasTextFrame** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape does not have a text frame and therefore cannot contain text.|
|**msoTrue**| The specified shape has a text frame and can therefore contain text.|

## Example

This example extracts text from all shapes on the first slide that contain text frames, and then it stores the names of these shapes and the text they contain in an array.


```vb
Dim shpTextArray() As Variant
Dim numShapes, numAutoShapes, i As Long

Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    numShapes = .Count
    If numShapes > 1 Then
        numTextShapes = 0
        ReDim shpTextArray(1 To 2, 1 To numShapes)
        For i = 1 To numShapes
            If .Item(i).HasTextFrame Then
                numTextShapes = numTextShapes + 1
                shpTextArray(numTextShapes, 1) = .Item(i).Name
                shpTextArray(numTextShapes, 2) = .Item(i) _
                    .TextFrame.TextRange.Text
            End If
        Next
        ReDim Preserve shpTextArray(1 To 2, 1 To numTextShapes)
    End If
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

