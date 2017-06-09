---
title: ShapeRange Object (Word)
keywords: vbawd10.chm2485
f1_keywords:
- vbawd10.chm2485
ms.prod: word
ms.assetid: 7112acc0-e241-16ef-77bc-101b72d05af0
ms.date: 06/08/2017
---


# ShapeRange Object (Word)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. 


## Remarks

You can include whichever shapes you want — chosen from among all the shapes in the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes in a document, all the selected shapes in a document, or all the freeform shapes in a document.Most operations that you can do with a **Shape** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, will cause an error.

Use  **Range** (Index), where Index is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use Visual Basic's **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on the active document.




```
ActiveDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```

The following example selects the shapes named "Oval 4" and "Rectangle 5" on the active document.




```
ActiveDocument.Shapes.Range(Array("Oval 4", "Rectangle 5")).Select
```

Although you can use the  **Range** method to return any number of shapes, it is simpler to use the **Item** method if you want to return only a single member of the collection. For example, `Shapes(1)`is simpler than `Shapes.Range(1)`.

Use  **ShapeRange** (Index), where Index is the name or the index number, to return a **Shape** object that represents a shape within a selection. The following example sets the fill for the first shape in the selection, assuming that the selection contains at least one shape.




```
Selection.ShapeRange(1).Fill.ForeColor.RGB = RGB(255, 0, 0)
```

This example selects all the shapes in the first section of the active document.




```
Set myRange = ActiveDocument.Sections(1).Range 
myRange.ShapeRange.Select
```

Use the  **Align**, **Distribute**, or **ZOrder** method to position a set of shapes relative to each other or relative to the document.

Use the  **Group**, **Regroup**, or **UnGroup** method to create and work with a single shape formed from a shape range. The **GroupItems** property for a **Shape** object returns the **GroupShapes** object, which represents all the shapes that were grouped to form one shape.

The recorder always uses the  **ShapeRange** property when recording shapes.


 **Note**  A  **ShapeRange** object doesn't include **InlineShape** objects.


## Methods



|**Name**|
|:-----|
|[Align](shaperange-align-method-word.md)|
|[Apply](shaperange-apply-method-word.md)|
|[CanvasCropBottom](shaperange-canvascropbottom-method-word.md)|
|[CanvasCropLeft](shaperange-canvascropleft-method-word.md)|
|[CanvasCropRight](shaperange-canvascropright-method-word.md)|
|[CanvasCropTop](shaperange-canvascroptop-method-word.md)|
|[ConvertToInlineShape](shaperange-converttoinlineshape-method-word.md)|
|[Delete](shaperange-delete-method-word.md)|
|[Distribute](shaperange-distribute-method-word.md)|
|[Duplicate](shaperange-duplicate-method-word.md)|
|[Flip](shaperange-flip-method-word.md)|
|[Group](shaperange-group-method-word.md)|
|[IncrementLeft](shaperange-incrementleft-method-word.md)|
|[IncrementRotation](shaperange-incrementrotation-method-word.md)|
|[IncrementTop](shaperange-incrementtop-method-word.md)|
|[Item](shaperange-item-method-word.md)|
|[PickUp](shaperange-pickup-method-word.md)|
|[ScaleHeight](shaperange-scaleheight-method-word.md)|
|[ScaleWidth](shaperange-scalewidth-method-word.md)|
|[Select](shaperange-select-method-word.md)|
|[SetShapesDefaultProperties](shaperange-setshapesdefaultproperties-method-word.md)|
|[Ungroup](shaperange-ungroup-method-word.md)|
|[ZOrder](shaperange-zorder-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Adjustments](shaperange-adjustments-property-word.md)|
|[AlternativeText](shaperange-alternativetext-property-word.md)|
|[Anchor](shaperange-anchor-property-word.md)|
|[Application](shaperange-application-property-word.md)|
|[AutoShapeType](shaperange-autoshapetype-property-word.md)|
|[BackgroundStyle](shaperange-backgroundstyle-property-word.md)|
|[Callout](shaperange-callout-property-word.md)|
|[CanvasItems](shaperange-canvasitems-property-word.md)|
|[Child](shaperange-child-property-word.md)|
|[Count](shaperange-count-property-word.md)|
|[Creator](shaperange-creator-property-word.md)|
|[Fill](shaperange-fill-property-word.md)|
|[Glow](shaperange-glow-property-word.md)|
|[GroupItems](shaperange-groupitems-property-word.md)|
|[Height](shaperange-height-property-word.md)|
|[HeightRelative](shaperange-heightrelative-property-word.md)|
|[HorizontalFlip](shaperange-horizontalflip-property-word.md)|
|[Hyperlink](shaperange-hyperlink-property-word.md)|
|[ID](shaperange-id-property-word.md)|
|[LayoutInCell](shaperange-layoutincell-property-word.md)|
|[Left](shaperange-left-property-word.md)|
|[LeftRelative](shaperange-leftrelative-property-word.md)|
|[Line](shaperange-line-property-word.md)|
|[LockAnchor](shaperange-lockanchor-property-word.md)|
|[LockAspectRatio](shaperange-lockaspectratio-property-word.md)|
|[Name](shaperange-name-property-word.md)|
|[Nodes](shaperange-nodes-property-word.md)|
|[Parent](shaperange-parent-property-word.md)|
|[ParentGroup](shaperange-parentgroup-property-word.md)|
|[PictureFormat](shaperange-pictureformat-property-word.md)|
|[Reflection](shaperange-reflection-property-word.md)|
|[RelativeHorizontalPosition](shaperange-relativehorizontalposition-property-word.md)|
|[RelativeHorizontalSize](shaperange-relativehorizontalsize-property-word.md)|
|[RelativeVerticalPosition](shaperange-relativeverticalposition-property-word.md)|
|[RelativeVerticalSize](shaperange-relativeverticalsize-property-word.md)|
|[Rotation](shaperange-rotation-property-word.md)|
|[Shadow](shaperange-shadow-property-word.md)|
|[ShapeStyle](shaperange-shapestyle-property-word.md)|
|[SoftEdge](shaperange-softedge-property-word.md)|
|[TextEffect](shaperange-texteffect-property-word.md)|
|[TextFrame](shaperange-textframe-property-word.md)|
|[TextFrame2](shaperange-textframe2-property-word.md)|
|[ThreeD](shaperange-threed-property-word.md)|
|[Title](shaperange-title-property-word.md)|
|[Top](shaperange-top-property-word.md)|
|[TopRelative](shaperange-toprelative-property-word.md)|
|[Type](shaperange-type-property-word.md)|
|[VerticalFlip](shaperange-verticalflip-property-word.md)|
|[Vertices](shaperange-vertices-property-word.md)|
|[Visible](shaperange-visible-property-word.md)|
|[Width](shaperange-width-property-word.md)|
|[WidthRelative](shaperange-widthrelative-property-word.md)|
|[WrapFormat](shaperange-wrapformat-property-word.md)|
|[ZOrderPosition](shaperange-zorderposition-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
