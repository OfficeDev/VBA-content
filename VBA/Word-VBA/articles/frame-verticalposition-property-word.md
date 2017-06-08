---
title: Frame.VerticalPosition Property (Word)
keywords: vbawd10.chm153747466
f1_keywords:
- vbawd10.chm153747466
ms.prod: word
api_name:
- Word.Frame.VerticalPosition
ms.assetid: 584880c0-85e3-d96c-291f-5671b792f818
ms.date: 06/08/2017
---


# Frame.VerticalPosition Property (Word)

Returns or sets the vertical distance between the edge of the frame and the item specified by the  **RelativeVerticalPosition** property. Read/write **Single** .


## Syntax

 _expression_ . **VerticalPosition**

 _expression_ Required. A variable that represents a **[Frame](frame-object-word.md)** object.


## Remarks

Can be a number that indicates a measurement in points, or can be any valid  **[WdFramePosition](wdframeposition-enumeration-word.md)** constant.


## Example

This example vertically aligns the first frame in the active document with the top of the page.


```vb
Set myFrame = ActiveDocument.Frames(1) 
With myFrame 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .VerticalPosition = wdFrameTop 
End With
```

This example adds a frame around the first shape in the active document and positions the frame 1 inch from the top margin.




```vb
If ActiveDocument.Shapes.Count >= 1 Then 
 ActiveDocument.Shapes(1).Select 
 Set aFrame = ActiveDocument.Frames.Add(Range:=Selection.Range) 
 With aFrame 
 .RelativeVerticalPosition = _ 
 wdRelativeVerticalPositionMargin 
 .VerticalPosition = InchesToPoints(1) 
 End With 
End If
```

This example vertically aligns the first table in the active document with the top of the page.




```vb
Set myTable = ActiveDocument.Tables(1).Rows 
With myTable 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .VerticalPosition = wdTableTop 
End With
```


## See also


#### Concepts


[Frame Object](frame-object-word.md)

