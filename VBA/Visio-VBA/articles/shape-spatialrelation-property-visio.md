---
title: Shape.SpatialRelation Property (Visio)
keywords: vis_sdr.chm11214400
f1_keywords:
- vis_sdr.chm11214400
ms.prod: visio
api_name:
- Visio.Shape.SpatialRelation
ms.assetid: 7e9f26b5-2887-493f-01c1-5e3900ea8c05
ms.date: 06/08/2017
---


# Shape.SpatialRelation Property (Visio)

Returns an integer that represents the spatial relationship of one shape to another shape. Both shapes must be on the same page or in the same master. Read-only.


## Syntax

 _expression_ . **SpatialRelation**( **_OtherShape_** , **_Tolerance_** , **_Flags_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OtherShape_|Required| **[IVSHAPE]**|The other  **Shape** object involved in the comparison.|
| _Tolerance_|Required| **Double**|A distance in internal drawing units with respect to the coordinate space defined by the  **Shape** object's parent.|
| _Flags_|Required| **Integer**|Flags that influence the result. See Remarks for the values of this argument.|

### Return Value

Integer


## Remarks


- The integer returned can be any combination of the values defined in  **[VisSpatialRelationCodes](visspatialrelationcodes-enumeration-visio.md)** in the Visio type library. The **SpatialRelation** property returns zero (0) if the two shapes being compared are not in any of the relationships discussed in the table in the **[VisSpatialRelationCodes](visspatialrelationcodes-enumeration-visio.md)** topic.
    
- The Flags argument can be any combination of the values of the constants defined in the following table. These constants are declared in  **VisSpatialRelationFlags** in the Visio type library. Use the NoShow cell to determine whether a Geometry section is hidden or visible. Hidden Geometry sections have a value of TRUE and visible Geometry sections have a value of FALSE in the NoShow cell.
    

|** Constant**|** Value**|** Description**|
|:-----|:-----|:-----|
| **visSpatialIncludeContainerShapes**|&;H80|Include containers. By default, containers are not included.|
| **visSpatialIncludeDataGraphics**|&;H40|Includes data graphic callout shapes and their sub-shapes. By default, data graphic callout shapes and their subshapes are not included. If the parent shape is itself a data graphic callout, searches are made between the parent shape's geometry and non-callout shapes, unless this flag is set.|
| **visSpatialIncludeGuides**|&;H2|Considers a guide's Geometry section. By default, guides do not influence the result.|
| **visSpatialIncludeHidden**|&;H10|Reserved for future use. Do not use.|
| **visSpatialIgnoreVisible**|&;H20|Does not consider visible Geometry sections. By default, visible Geometry sections influence the result.|

 **Note**   When it compares two shapes, the **SpatialRelation** property does not consider the width of a shape's line, shadows, line ends, control points, or connection points.


## Example

This Microsoft Visual Basic for Applications (VBA) example shows how to use the  **SpatialRelation** property in an event handler for the **ShapeAdded** event to determine the spatial relationship between shapes.

Before adding the following code to your VBA project, make sure there is at least one shape on the drawing page. Then, after adding the code, add another shape to your drawing.




```vb
Public Sub Document_ShapeAdded(ByVal Shape As IVShape) 
 
    Dim vsoShapeOnPage As Visio.Shape  
    Dim intTolerance As Integer 
    Dim intReturnValue As VisSpatialRelationCodes  
    Dim intFlag As VisSpatialRelationFlags  
    Dim strReturn As String 
    On Error GoTo errHandler  
 
    'Initialize tolerance argument. 
    intTolerance = 0.25  
 
    'Initialize flags argument. 
    intFlag = visSpatialIncludeHidden  
    For Each vsoShapeOnPage In ActivePage.Shapes  
 
        'Get the spatial relationship. 
        intReturnValue = Shape.SpatialRelation(vsoShapeOnPage, _  
            intTolerance, intFlag)  
 
        'Convert return code to string value. 
        Select Case intReturnValue       
            Case VisSpatialRelationCodes.visSpatialContain  
                strReturn = "Contains"  
            Case VisSpatialRelationCodes.visSpatialContainedIn  
                strReturn = "is Contained in"  
            Case VisSpatialRelationCodes.visSpatialOverlap  
                strReturn = "overlaps"  
            Case VisSpatialRelationCodes.visSpatialTouching  
                strReturn = "is touching"  
            Case Else 
                strReturn = "has no relation with"  
        End Select  
        
        'Display relationship in the shape's text. 
        vsoShapeOnPage.Text = Shape.Name &; " " &; strReturn &; " " &; _  
            vsoShapeOnPage.Name  
 
    Next  
 
errHandler:  
 
End Sub
```


