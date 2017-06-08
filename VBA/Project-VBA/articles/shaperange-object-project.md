---
title: ShapeRange Object (Project)
ms.prod: project-server
ms.assetid: 315031aa-4b8c-424b-26e7-ce15897beb05
ms.date: 06/08/2017
---


# ShapeRange Object (Project)
Represents a shape range, which is a collection of one or more shapes in a report.
 

## Remarks

Project uses the same Office Art infrastructure that other Office applications use, and adapts Office Art to reports, tables, and charts that can use fields in the active project. However, Project does not implement all  **ShapeRange** operations. For example, Project does not support automatic alignment, distribution, grouping, or merging of shapes in a shape range.
 

 
A shape range can contain a single shape or all the shapes in the report. You can include whichever shapes you want to construct a shape range. For example, you can construct a  **ShapeRange** collection that contains the first three shapes in a report, all the shapes in a report, or only the triangle shapes.
 

 
Most operations that you can do with a  **Shape** object, you can also do with a **ShapeRange** object that contains only one shape. Some operations, when performed on a **ShapeRange** object that contains more than one shape, shapes of different types, or a shape that is not fully supported in Project, can cause an error. For example, if a shape range contains a rectangle and a chart, and you try to set the **Fill** property, the statement fails because a chart does not implement the **Fill** property. In other cases, for example if you use the **Rotation** property on a shape range that contains a chart and a rectangle, Project rotates the rectangle but silently ignores the chart.
 

 

## Examples

You can return a set of shapes that are specified by the index number or by the shape name. Use  `Shapes.Range(index)`, where  _index_ is an array of index numbers or names. For example, both of the following statements are valid:
 

 

```
Set myRange1 = theReport.Shapes.Range(Array(1, 2))
Set myRange2 = theReport.Shapes.Range(Array("Textbox 1", "Textbox 2"))
```

To create a  **ShapeRange** object that contains all of the shapes in the report, use a statement such as the following:
 

 



```
Set allShapes = theReport.Shapes.Range(Array(1, theReport.Shapes.Count))
```

To create a  **ShapeRange** object with a single member of the **Shapes** collection, you can use statements such as the following:
 

 



```
Set myRange3 = theReport.Shapes.Range(2)
Set myRange4 = theReport.Shapes.Range("Rectangle 2")
```

To perform an operation on a single shape in a  **ShapeRange** collection, you can use statements such as the following:
 

 



```
myRange1(2).Fill.ForeColor.RGB = RGB(120, 120, 80)
myRange1("Textbox 2").Fill.ForeColor.RGB = RGB(120, 120, 80)
```

Alternately, you can perform an operation directly on a  **Shape** object, without using a shape range.
 

 



```
theReport.Shapes("Big rectangle").Fill.ForeColor.RGB = RGB(120, 120, 80)
```


## Methods



|**Description**|
|:-----|
|The  **Align** method is not implemented in Project.|
|Applies formatting to a shape range, where the formatting information has been copied by using the  **[PickUp](shape-pickup-method-project.md)** method.|
|Copies the shape range to the Clipboard.|
|Cuts the shape range to the Clipboard.|
|Deletes the shape range.|
|The  **Distribute** method is not implemented in Project.|
|Duplicates a shape range and returns a reference to the copy.|
|Flips each shape in the shape range around its horizontal or vertical axis.|
|The  **Group** method is not implemented in Project.|
|Moves each shape in the shape range horizontally by the specified number of points.|
|Rotates each shape in the shape range around the z-axis by the specified number of degrees.|
|Moves each shape in the shape range vertically by the specified number of points.|
|Gets an individual  **Shape** object in the shape range collection.|
|The  **MergeShapes** method is not implemented in Project.|
|Copies the formatting of the shape range.|
|The  **Regroup** method is not implemented in Project.|
|The  **RerouteConnections** method is not implemented in Project.|
|Scales the height of the range of shapes by a specified factor.|
|Scales the width of the range of shapes by a specified factor.|
|Selects each shape in a shape range.|
|Applies the formatting of a default shape to each shape in the range.|
|The  **Ungroup** method is not implemented in Project.|
|Moves the shape range in front of or behind other shapes (that is, changes the position in the z-order).|

## Properties



|**Name**|
|:-----|
|[Adjustments](shaperange-adjustments-property-project.md)|
|[AlternativeText](shaperange-alternativetext-property-project.md)|
|[Application](shaperange-application-property-project.md)|
|[AutoShapeType](shaperange-autoshapetype-property-project.md)|
|[BackgroundStyle](shaperange-backgroundstyle-property-project.md)|
|[BlackWhiteMode](shaperange-blackwhitemode-property-project.md)|
|[Callout](shaperange-callout-property-project.md)|
|[Chart](shaperange-chart-property-project.md)|
|[Child](shaperange-child-property-project.md)|
|[ConnectionSiteCount](shaperange-connectionsitecount-property-project.md)|
|[Connector](shaperange-connector-property-project.md)|
|[ConnectorFormat](shaperange-connectorformat-property-project.md)|
|[Count](shaperange-count-property-project.md)|
|[Fill](shaperange-fill-property-project.md)|
|[Glow](shaperange-glow-property-project.md)|
|[GroupItems](shaperange-groupitems-property-project.md)|
|[HasChart](shaperange-haschart-property-project.md)|
|[HasTable](shaperange-hastable-property-project.md)|
|[Height](shaperange-height-property-project.md)|
|[HorizontalFlip](shaperange-horizontalflip-property-project.md)|
|[ID](shaperange-id-property-project.md)|
|[Left](shaperange-left-property-project.md)|
|[Line](shaperange-line-property-project.md)|
|[LockAspectRatio](shaperange-lockaspectratio-property-project.md)|
|[Name](shaperange-name-property-project.md)|
|[Nodes](shaperange-nodes-property-project.md)|
|[Parent](shaperange-parent-property-project.md)|
|[ParentGroup](shaperange-parentgroup-property-project.md)|
|[Reflection](shaperange-reflection-property-project.md)|
|[Rotation](shaperange-rotation-property-project.md)|
|[Script](shaperange-script-property-project.md)|
|[Shadow](shaperange-shadow-property-project.md)|
|[ShapeStyle](shaperange-shapestyle-property-project.md)|
|[SoftEdge](shaperange-softedge-property-project.md)|
|[Table](shaperange-table-property-project.md)|
|[TextEffect](shaperange-texteffect-property-project.md)|
|[TextFrame](shaperange-textframe-property-project.md)|
|[TextFrame2](shaperange-textframe2-property-project.md)|
|[ThreeD](shaperange-threed-property-project.md)|
|[Title](shaperange-title-property-project.md)|
|[Top](shaperange-top-property-project.md)|
|[Type](shaperange-type-property-project.md)|
|[Value](shaperange-value-property-project.md)|
|[VerticalFlip](shaperange-verticalflip-property-project.md)|
|[Vertices](shaperange-vertices-property-project.md)|
|[Visible](shaperange-visible-property-project.md)|
|[Width](shaperange-width-property-project.md)|
|[ZOrderPosition](shaperange-zorderposition-property-project.md)|

## See also


#### Other resources


 
[Shapes Object](shapes-object-project.md)
 
[Shape Object](shape-object-project.md)
