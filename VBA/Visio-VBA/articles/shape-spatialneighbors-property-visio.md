---
title: Shape.SpatialNeighbors Property (Visio)
keywords: vis_sdr.chm11214395
f1_keywords:
- vis_sdr.chm11214395
ms.prod: visio
api_name:
- Visio.Shape.SpatialNeighbors
ms.assetid: 98069519-d788-c34f-ac25-64bda73324d5
ms.date: 06/08/2017
---


# Shape.SpatialNeighbors Property (Visio)

Returns a  **Selection** object that represents the shapes that meet certain criteria in relation to a specified shape. Read-only.


## Syntax

 _expression_ . **SpatialNeighbors**( **_Relation_** , **_Tolerance_** , **_Flags_** , **_ResultRoot_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Relation_|Required| **Integer**|An integer describing the type of relationship to be used.|
| _Tolerance_|Required| **Double**|A distance in internal drawing units with respect to the coordinate space defined by the parent shape.|
| _Flags_|Required| **Integer**|Flags that influence the type of entries returned in results.|
| _ResultRoot_|Optional| **Variant**|A  **Shape** object that represents a page or group.|

### Return Value

Selection


## Remarks

For values of the  _Relation_ argument, see the **[SpatialRelation](shape-spatialrelation-property-visio.md)** property topic.

The  _Flags_ argument can be any combination of the values of the constants defined in the following table. These constants are also defined in **VisSpatialRelationFlags** in the Visio type library.



|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visSpatialIncludeContainerShapes**|&;H80|Include containers. By default, containers are not included.|
| **visSpatialIncludeDataGraphics**|&;H40|Includes data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.|
| **visSpatialIncludeGuides**| &;H2| Considers a guide's Geometry section. By default, guides do not influence the result.|
| **visSpatialFrontToBack**| &;H4| Orders items front to back.|
| **visSpatialBackToFront**| &;H8| Orders items back to front.|
| **visSpatialIncludeHidden**| &;H10| Considers hidden Geometry sections. By default, hidden Geometry sections do not influence the result.|
| **visSpatialIgnoreVisible**| &;H20| Does not consider visible Geometry sections. By default, visible Geometry sections influence the result.|
Use the NoShow cell to determine whether a Geometry section is hidden or visible. Hidden Geometry sections have a value of TRUE and visible Geometry sections have a value of FALSE in the NoShow cell.

If  _Relation_ is not specified, the **SpatialNeighbors** property uses all the possible relationships as criteria.

Beginning with Visio 2002, if  _Flags_ contains **VisSpatialFrontToBack** , items in the **Selection** object returned by the **SpatialNeighbors** property are ordered front to back. If **visSpatialBackToFront** is set, the items returned are ordered back to front. If this flag is not set, or if you are running an earlier version of Visio, the order is unpredictable. You can determine the order by using the **Index** property of the shapes identified in the **Selection** object.

If you don't specify  _ResultRoot_, this property returns a  **Selection** object that represents the shapes that meet certain criteria in relation to the specified shape. If you specify _ResultRoot_, this property returns a  **Selection** object that represents all the shapes in the **Shape** object specified by _ResultRoot_ that meet certain criteria in relation to the specified shape. For example, specify _ResultRoot_ to find all shapes within a group that are near a specified shape.

If  _ResultRoot_ is specified but isn't on the same page or in the same master as the **Shape** object to which you are comparing it, the **SpatialNeighbors** property raises an exception and returns **Nothing** .

When it compares two shapes, the  **SpatialNeighbors** property does not consider the width of a shape's line, shadows, line ends, control points, or connection points.


## Example

This Microsoft Visual Basic for Applications (VBA) example shows how to use the  **SpatialNeighbors** property in an event handler for the **ShapeAdded** event to determine if one shape is contained within another.

Before adding the following code to your VBA project, add at least one shape to your drawing. Then add another shape to your drawing, either by dragging it from a stencil or drawing it, positioning it so that it is completely contained within an existing shape.




```vb
 
Public Sub Document_ShapeAdded(ByVal Shape As IVShape) 
 
 Dim vsoShapeOnPage As Visio.Shape 
 Dim intTolerance As Integer 
 Dim vsoReturnedSelection As Visio.Selection 
 Dim strSpatialRelation As String 
 Dim intSpatialRelation As VisSpatialRelationCodes 
 
 On Error GoTo errHandler 
 
 'Initialize string 
 strSpatialRelation = "" 
 
 'Set tolerance argument 
 intTolerance = 0.25 
 
 'Set Spatial Relation argument 
 intSpatialRelation = visSpatialContainedIn 
 
 'Get the set of spatially related shapes 
 'that meet the criteria set by the arguments. 
 Set vsoReturnedSelection = Shape.SpatialNeighbors _ 
 (intSpatialRelation, intTolerance, 0) 
 
 'Evaluate the results. 
 If vsoReturnedSelection.Count = 0 Then 
 
 'No shapes met the criteria set by 
 'the arguments of the method. 
 strSpatialRelation = Shape.Name &; " is not contained." 
 
 Else 
 
 'Build the positive result string. 
 For Each vsoShapeOnPage In vsoReturnedSelection 
 strSpatialRelation = strSpatialRelation &; _ 
 Shape.Name &; " is contained by " &; _ 
 vsoShapeOnPage.Name &; Chr$(10) 
 
 Next 
 
 End If 
 
 'Display the results on the shape added. 
 Shape.Text = strSpatialRelation 
 
 errHandler: 
 
End Sub
```


