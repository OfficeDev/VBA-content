---
title: Shape Object (Excel)
keywords: vbaxl10.chm635072
f1_keywords:
- vbaxl10.chm635072
ms.prod: excel
api_name:
- Excel.Shape
ms.assetid: 8f01fcd1-b7d9-5216-2de5-40fb6648a403
ms.date: 06/08/2017
---


# Shape Object (Excel)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks

 The **Shape** object is a member of the **[Shapes](shapes-object-excel.md)** collection. The **Shapes** collection contains all the shapes in a workbook.


 **Note**  There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a workbook; the **[ShapeRange](shaperange-object-excel.md)** collection, which represents a specified subset of the shapes on a workbook (for example, a **ShapeRange** object could represent shapes one and four in the workbook, or it could represent all the selected shapes in the workbook); and the **Shape** object, which represents a single shape on a worksheet. If you want to work with several shapes at the same time or with shapes within the selection, use a **ShapeRange** collection.


### Using the Shape Object

The following sections describes how to:


- Return the shapes attached to the ends of a connector.
    
- Return a newly created freeform.
    
- Return a single shape from within a group.
    
- Return a newly formed group of shapes.
    
- Return an existing shape.
    
- Return a shape within the selection.
    

### Returning the Shapes Attached to the Ends of a Connector

To return a  **Shape** object that represents one of the shapes attached by a connector, use the **[BeginConnectedShape](connectorformat-beginconnectedshape-property-excel.md)** or **[EndConnectedShape](connectorformat-endconnectedshape-property-excel.md)** property.


### Returning a newly created freeform

Use the  **[BuildFreeform](shapes-buildfreeform-method-excel.md)** and **[AddNodes](freeformbuilder-addnodes-method-excel.md)** methods to define the geometry of a new freeform, and use the **[ConvertToShape](freeformbuilder-converttoshape-method-excel.md)** method to create the freeform and return the **Shape** object that represents it.


### Returning a Single Shape from Within a Group

Use  **[GroupItems](shape-groupitems-property-excel.md)** ( _index_ ), where _index_ is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape.


### Returning a Newly Formed Group of Shapes

Use the  **[Group](shaperange-group-method-excel.md)** or **[Regroup](shaperange-regroup-method-excel.md)** method to group a range of shapes and return a single **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape.


### Returning an Existing Shape

Use  **[Shapes](worksheet-shapes-property-excel.md)** ( _index_ ), where _index_ is the shape name or the index number, to return a **Shape** object that represents a shape.


### Returning a Shape Within the Selection

Use  `Selection.ShapeRange`( _index_ ), where _index_ is the shape name or the index number, to return a **Shape** object that represents a shape within the selection.


## Example

The following example horizontally flips shape one and the shape named Rectangle 1 on  _myDocument_.


```
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).Flip msoFlipHorizontal 
myDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

Each shape is assigned a default name when you add it to the  **Shapes** collection. To give the shape a more meaningful name, use the **Name** property. The following example adds a rectangle to myDocument, gives it the name Red Square, and then sets its foreground color and line style.




```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 144, 144, 72, 72) 
 .Name = "Red Square" 
 .Fill.ForeColor.RGB = RGB(255, 0, 0) 
 .Line.DashStyle = msoLineDashDot 
End With
```

The following example sets the fill for the first shape in the selection in the active window, assuming that there's at least one shape in the selection.




```
ActiveWindow.Selection.ShapeRange(1).Fill.ForeColor.RGB = _ 
 RGB(255, 0, 0)
```


## Methods



|**Name**|
|:-----|
|[Apply](shape-apply-method-excel.md)|
|[Copy](shape-copy-method-excel.md)|
|[CopyPicture](shape-copypicture-method-excel.md)|
|[Cut](shape-cut-method-excel.md)|
|[Delete](shape-delete-method-excel.md)|
|[Duplicate](shape-duplicate-method-excel.md)|
|[Flip](shape-flip-method-excel.md)|
|[IncrementLeft](shape-incrementleft-method-excel.md)|
|[IncrementRotation](shape-incrementrotation-method-excel.md)|
|[IncrementTop](shape-incrementtop-method-excel.md)|
|[PickUp](shape-pickup-method-excel.md)|
|[RerouteConnections](shape-rerouteconnections-method-excel.md)|
|[ScaleHeight](shape-scaleheight-method-excel.md)|
|[ScaleWidth](shape-scalewidth-method-excel.md)|
|[Select](shape-select-method-excel.md)|
|[SetShapesDefaultProperties](shape-setshapesdefaultproperties-method-excel.md)|
|[Ungroup](shape-ungroup-method-excel.md)|
|[ZOrder](shape-zorder-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Adjustments](shape-adjustments-property-excel.md)|
|[AlternativeText](shape-alternativetext-property-excel.md)|
|[Application](shape-application-property-excel.md)|
|[AutoShapeType](shape-autoshapetype-property-excel.md)|
|[BackgroundStyle](shape-backgroundstyle-property-excel.md)|
|[BlackWhiteMode](shape-blackwhitemode-property-excel.md)|
|[BottomRightCell](shape-bottomrightcell-property-excel.md)|
|[Callout](shape-callout-property-excel.md)|
|[Chart](shape-chart-property-excel.md)|
|[Child](shape-child-property-excel.md)|
|[ConnectionSiteCount](shape-connectionsitecount-property-excel.md)|
|[Connector](shape-connector-property-excel.md)|
|[ConnectorFormat](shape-connectorformat-property-excel.md)|
|[ControlFormat](shape-controlformat-property-excel.md)|
|[Creator](shape-creator-property-excel.md)|
|[Fill](shape-fill-property-excel.md)|
|[FormControlType](shape-formcontroltype-property-excel.md)|
|[Glow](shape-glow-property-excel.md)|
|[GroupItems](shape-groupitems-property-excel.md)|
|[HasChart](shape-haschart-property-excel.md)|
|[HasSmartArt](shape-hassmartart-property-excel.md)|
|[Height](shape-height-property-excel.md)|
|[HorizontalFlip](shape-horizontalflip-property-excel.md)|
|[Hyperlink](shape-hyperlink-property-excel.md)|
|[ID](shape-id-property-excel.md)|
|[Left](shape-left-property-excel.md)|
|[Line](shape-line-property-excel.md)|
|[LinkFormat](shape-linkformat-property-excel.md)|
|[LockAspectRatio](shape-lockaspectratio-property-excel.md)|
|[Locked](shape-locked-property-excel.md)|
|[Name](shape-name-property-excel.md)|
|[Nodes](shape-nodes-property-excel.md)|
|[OLEFormat](shape-oleformat-property-excel.md)|
|[OnAction](shape-onaction-property-excel.md)|
|[Parent](shape-parent-property-excel.md)|
|[ParentGroup](shape-parentgroup-property-excel.md)|
|[PictureFormat](shape-pictureformat-property-excel.md)|
|[Placement](shape-placement-property-excel.md)|
|[Reflection](shape-reflection-property-excel.md)|
|[Rotation](shape-rotation-property-excel.md)|
|[Shadow](shape-shadow-property-excel.md)|
|[ShapeStyle](shape-shapestyle-property-excel.md)|
|[SmartArt](shape-smartart-property-excel.md)|
|[SoftEdge](shape-softedge-property-excel.md)|
|[TextEffect](shape-texteffect-property-excel.md)|
|[TextFrame](shape-textframe-property-excel.md)|
|[TextFrame2](shape-textframe2-property-excel.md)|
|[ThreeD](shape-threed-property-excel.md)|
|[Title](shape-title-property-excel.md)|
|[Top](shape-top-property-excel.md)|
|[TopLeftCell](shape-topleftcell-property-excel.md)|
|[Type](shape-type-property-excel.md)|
|[VerticalFlip](shape-verticalflip-property-excel.md)|
|[Vertices](shape-vertices-property-excel.md)|
|[Visible](shape-visible-property-excel.md)|
|[Width](shape-width-property-excel.md)|
|[ZOrderPosition](shape-zorderposition-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
