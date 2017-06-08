---
title: Selection Object (PowerPoint)
keywords: vbapp10.chm508000
f1_keywords:
- vbapp10.chm508000
ms.prod: powerpoint
api_name:
- PowerPoint.Selection
ms.assetid: a7def3bd-9dff-da53-152d-4fd686642413
ms.date: 06/08/2017
---


# Selection Object (PowerPoint)

Represents the selection in the specified document window. The  **Selection** object is deleted whenever you change slides in an active slide view (the **Type** property will return **ppSelectionNone** ).


## Example

Use the [Selection](http://msdn.microsoft.com/library/3773ff08-043d-2b57-25ea-ba44cc30c77a%28Office.15%29.aspx)property to return the  **Selection** object. The following example places a copy of the selection in the active window on the Clipboard.


```
ActiveWindow.Selection.Copy
```

Use the [ShapeRange](http://msdn.microsoft.com/library/3fd7aed0-ab63-adaa-1a46-c745b6c3e245%28Office.15%29.aspx), [SlideRange](http://msdn.microsoft.com/library/2d853875-b0c2-ab8e-38b6-4e1397d4e669%28Office.15%29.aspx), or [TextRange](http://msdn.microsoft.com/library/532c0a35-c18d-8030-8e6a-3f1cdb47c244%28Office.15%29.aspx)property to return a range of shapes, slides, or text from the selection.

The following example sets the fill foreground color for the selected shapes in window two, assuming that there's at least one shape selected, and assuming that all selected shapes have a fill whose forecolor can be set.




```
With Windows(2).Selection.ShapeRange.Fill

    .Visible = True

    .ForeColor.RGB = RGB(255, 0, 255)

End With
```

The following example sets the text in the first selected shape in window two if that shape contains a text frame.




```
With Windows(2).Selection.ShapeRange(1)

    If .HasTextFrame Then

        .TextFrame.TextRange = "Current Choice"

    End If

End With
```

The following example cuts the selected text in the active window and places it on the Clipboard.




```
ActiveWindow.Selection.TextRange.Cut
```

The following example duplicates all the slides in the selection (if you're in slide view, this duplicates the current slide).




```
ActiveWindow.Selection.SlideRange.Duplicate
```

If you don't have an object of the appropriate type selected when you use one of these properties (for instance, if you use the  **ShapeRange** property when there are no shapes selected), an error occurs. Use the[Type](http://msdn.microsoft.com/library/1c39388f-2ca4-211c-393e-1f0af0723898%28Office.15%29.aspx)property to determine what kind of object or objects are selected. The following example checks to see whether the selection contains slides. If the selection does contain slides, the example sets the background for the first slide in the selection.




```
With Windows(2).Selection

    If .Type = ppSelectionSlides Then

        With .SlideRange(1)

            .FollowMasterBackground = False

            .Background.Fill.PresetGradient _

                msoGradientHorizontal, 1, msoGradientLateSunset

        End With

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/954106da-a2a9-0c55-114a-5a79f578e0c4%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/305103ad-f4d1-8173-e331-17750587d865%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/879d15ca-97b4-cf44-27a0-7e15f6041b34%28Office.15%29.aspx)|
|[Unselect](http://msdn.microsoft.com/library/376a6b26-e877-c50c-c4ce-82273afc1fb8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/eb1591fe-f6ce-1f9c-21e1-fab39589c527%28Office.15%29.aspx)|
|[ChildShapeRange](http://msdn.microsoft.com/library/f7458e07-47ec-c832-0731-94f4ba94ca89%28Office.15%29.aspx)|
|[HasChildShapeRange](http://msdn.microsoft.com/library/f86dac76-66cc-7512-fe7c-1a16f5a381f8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/01f9d99a-0ace-4ec3-121b-e22c35240406%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/3fd7aed0-ab63-adaa-1a46-c745b6c3e245%28Office.15%29.aspx)|
|[SlideRange](http://msdn.microsoft.com/library/2d853875-b0c2-ab8e-38b6-4e1397d4e669%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/532c0a35-c18d-8030-8e6a-3f1cdb47c244%28Office.15%29.aspx)|
|[TextRange2](http://msdn.microsoft.com/library/3c4ccea8-d234-d7ab-a9d1-57b53b169066%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/1c39388f-2ca4-211c-393e-1f0af0723898%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
