---
title: ShapeRange Object (Excel)
keywords: vbaxl10.chm639072
f1_keywords:
- vbaxl10.chm639072
ms.prod: excel
api_name:
- Excel.ShapeRange
ms.assetid: e1b8229c-73a0-4a77-5e00-4bcec9032260
ms.date: 06/08/2017
---


# ShapeRange Object (Excel)

Represents a shape range, which is a set of shapes on a document.


## Remarks

 A shape range can contain as few as a single shape or as many as all the shapes on the document. You can include whichever shapes you want — chosen from among all the shapes on the document or all the shapes in the selection — to construct a shape range. For example, you could construct a **ShapeRange** collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.


## Example

 **Returning a Set of Shapes You Specify by Name or Index Number**



Use  `Shapes.Range`( _index_ ), where _index_ is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use the **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on _myDocument_.




```
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)).Fill.Patterned _ 
 msoPatternHorizontalBrick
```

The following example sets the fill pattern for the shapes named Oval 4 and Rectangle 5 on  _myDocument_.

Although you can use the  **Range** property to return any number of shapes or slides, it's simpler to use the **Item** method if you want to return only a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.




```
Set myDocument = Worksheets(1) 
Set myRange = myDocument.Shapes.Range(Array("Oval 4", _ 
 "Rectangle 5")) 
myRange.Fill.Patterned msoPatternHorizontalBrick
```

 **Returning All or Some of the Selected Shapes on a Document**

Use the  **ShapeRange** property of the **Selection** object to return all the shapes in the selection. The following example sets the fill foreground color for all the shapes in the selection in window one, assuming that there's at least one shape in the selection.




```
Windows(1).Selection.ShapeRange.Fill.ForeColor.RGB = _ 
 RGB(255, 0, 255)
```

Use  `Selection.ShapeRange`( _index_ ), where _index_ is the shape name or the index number, to return a single shape within the selection. The following example sets the fill foreground color for shape two in the collection of selected shapes in window one, assuming that there are at least two shapes in the selection.






```
Windows(1).Selection.ShapeRange(2).Fill.ForeColor.RGB = _ 
 RGB(255, 0, 255)
```


## Methods



|**Name**|
|:-----|
|[Align](shaperange-align-method-excel.md)|
|[Apply](shaperange-apply-method-excel.md)|
|[Delete](shaperange-delete-method-excel.md)|
|[Distribute](shaperange-distribute-method-excel.md)|
|[Duplicate](shaperange-duplicate-method-excel.md)|
|[Flip](shaperange-flip-method-excel.md)|
|[Group](shaperange-group-method-excel.md)|
|[IncrementLeft](shaperange-incrementleft-method-excel.md)|
|[IncrementRotation](shaperange-incrementrotation-method-excel.md)|
|[IncrementTop](shaperange-incrementtop-method-excel.md)|
|[Item](shaperange-item-method-excel.md)|
|[PickUp](shaperange-pickup-method-excel.md)|
|[Regroup](shaperange-regroup-method-excel.md)|
|[RerouteConnections](shaperange-rerouteconnections-method-excel.md)|
|[ScaleHeight](shaperange-scaleheight-method-excel.md)|
|[ScaleWidth](shaperange-scalewidth-method-excel.md)|
|[Select](shaperange-select-method-excel.md)|
|[SetShapesDefaultProperties](shaperange-setshapesdefaultproperties-method-excel.md)|
|[Ungroup](shaperange-ungroup-method-excel.md)|
|[ZOrder](shaperange-zorder-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Adjustments](shaperange-adjustments-property-excel.md)|
|[AlternativeText](shaperange-alternativetext-property-excel.md)|
|[Application](shaperange-application-property-excel.md)|
|[AutoShapeType](shaperange-autoshapetype-property-excel.md)|
|[BackgroundStyle](shaperange-backgroundstyle-property-excel.md)|
|[BlackWhiteMode](shaperange-blackwhitemode-property-excel.md)|
|[Callout](shaperange-callout-property-excel.md)|
|[Chart](shaperange-chart-property-excel.md)|
|[Child](shaperange-child-property-excel.md)|
|[ConnectionSiteCount](shaperange-connectionsitecount-property-excel.md)|
|[Connector](shaperange-connector-property-excel.md)|
|[ConnectorFormat](shaperange-connectorformat-property-excel.md)|
|[Count](shaperange-count-property-excel.md)|
|[Creator](shaperange-creator-property-excel.md)|
|[Fill](shaperange-fill-property-excel.md)|
|[Glow](shaperange-glow-property-excel.md)|
|[GroupItems](shaperange-groupitems-property-excel.md)|
|[HasChart](shaperange-haschart-property-excel.md)|
|[Height](shaperange-height-property-excel.md)|
|[HorizontalFlip](shaperange-horizontalflip-property-excel.md)|
|[ID](shaperange-id-property-excel.md)|
|[Left](shaperange-left-property-excel.md)|
|[Line](shaperange-line-property-excel.md)|
|[LockAspectRatio](shaperange-lockaspectratio-property-excel.md)|
|[Name](shaperange-name-property-excel.md)|
|[Nodes](shaperange-nodes-property-excel.md)|
|[Parent](shaperange-parent-property-excel.md)|
|[ParentGroup](shaperange-parentgroup-property-excel.md)|
|[PictureFormat](shaperange-pictureformat-property-excel.md)|
|[Reflection](shaperange-reflection-property-excel.md)|
|[Rotation](shaperange-rotation-property-excel.md)|
|[Shadow](shaperange-shadow-property-excel.md)|
|[ShapeStyle](shaperange-shapestyle-property-excel.md)|
|[SoftEdge](shaperange-softedge-property-excel.md)|
|[TextEffect](shaperange-texteffect-property-excel.md)|
|[TextFrame](shaperange-textframe-property-excel.md)|
|[TextFrame2](shaperange-textframe2-property-excel.md)|
|[ThreeD](shaperange-threed-property-excel.md)|
|[Title](shaperange-title-property-excel.md)|
|[Top](shaperange-top-property-excel.md)|
|[Type](shaperange-type-property-excel.md)|
|[VerticalFlip](shaperange-verticalflip-property-excel.md)|
|[Vertices](shaperange-vertices-property-excel.md)|
|[Visible](shaperange-visible-property-excel.md)|
|[Width](shaperange-width-property-excel.md)|
|[ZOrderPosition](shaperange-zorderposition-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
