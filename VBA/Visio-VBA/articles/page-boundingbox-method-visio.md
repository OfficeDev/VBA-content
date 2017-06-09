---
title: Page.BoundingBox Method (Visio)
keywords: vis_sdr.chm10916090
f1_keywords:
- vis_sdr.chm10916090
ms.prod: visio
api_name:
- Visio.Page.BoundingBox
ms.assetid: f281e304-057f-5555-8efd-fd81d088b8cd
ms.date: 06/08/2017
---


# Page.BoundingBox Method (Visio)

Returns a rectangle that tightly encloses the shapes of a page.


## Syntax

 _expression_ . **BoundingBox**( **_Flags_** , **_lpr8Left_** , **_lpr8Bottom_** , **_lpr8Right_** , **_lpr8Top_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Flags_|Required| **Integer**|Flags that influence the bounding box calculated for each shape that contributes to the resulting bounding box.|
| _lpr8Left_|Required| **Double**|Returns the x-coordinate of the left edge of the bounding box.|
| _lpr8Bottom_|Required| **Double**|Returns the y-coordinate of the bottom edge of the bounding box.|
| _lpr8Right_|Required| **Double**|Returns the x-coordinate of the right edge of the bounding box.|
| _lpr8Top_|Required| **Double**|Returns the y-coordinate of the top edge of the bounding box.|

### Return Value

Nothing


## Remarks

For a  **Shape** object, the **BoundingBox** method returns a rectangle that tightly encloses the shape and its subshapes.

For a  **Page** , **Master** , or **Selection** object, the **BoundingBox** method returns a rectangle that tightly encloses the page's, master's, or selection's shapes and their subshapes.

If the  **BoundingBox** method returns an error, or if it is asked to return the rectangle enclosing zero shapes, the rectangle returned is { left: 0, bottom: 0, right: -1, top: -1 }; otherwise, the rectangle returned has left less than or equal to (<=) right, and bottom less than or equal to (<=) top. The numbers returned are in internal units (inches).

The bounding rectangle returned for an individual shape depends on its  **Type** property.



|**Constant**|**Description**|
|:-----|:-----|
| **visTypePage**|Equivalent to  **Page.BoundingBox** or **Master.BoundingBox** .|
| **visTypeGroup**|Rectangle that tightly encloses the group and its subshapes.|
| **visTypeShape**|Determined rectangle depends on flags. See the following table.|
| **visTypeForeignObject**|Determined rectangle depends on flags. See the following table.|
| **visTypeGuide**|Determined rectangle depends on flags. See the following table.|
The method will raise an exception for object type  **visTypeDoc** .

The  _Flags_ argument has several bits that control the bounding box retrieved for each shape. If more than one of the bits described in the following table is set, the rectangle determined for the shape covers all rectangles implied by the bits.



|**Flag**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visBBoxUprightWH**|&;H1|Returns a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's width-height box.If the shape is not rotated, its upright width-height box and its width-height box are the same. Paths in the shape's geometry need not and often do not lie entirely within the shape's width-height box.|
| **visBBoxUprightText**|&;H2|Returns a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the shape's text.|
| **visBBoxExtents**|&;H4|Returns a rectangle that is the smallest rectangle parallel to the local coordinate system of the shape's parent that encloses the paths stroked by the shape's geometry.This may be larger or smaller than the shape's upright width-height box. The extents box determined for a shape of type  **visTypeForeignObject** equals that shape's upright width-height box.|
| **visBBoxIncludeHidden**|&;H10|Includes hidden geometry.|
| **visBBoxIgnoreVisible**|&;H20|Ignores visible geometry.|
| **visBBoxIncludeDataGraphics**|&;H10000|Includes data-graphic callout shapes (and their sub-shapes) that are applied to the shapes on the page. Off by default.|
| **visBBoxIncludeGuides**|&;H1000|Includes extents for shapes of type  **visTypeguide** . By default, the extents of shapes of type **visTypeGuide** are ignored.If you request guide extents, then only the _x_ positions of vertical guides and the _y_ positions of horizontal guides contribute to the rectangle that is returned. If any vertical guides are reported on, an infinite _y_ extent is returned. If any horizontal guides are reported on, an infinite _ x_ extent is returned. If any rotated guides are reported on, infinite _x_ and _y_ extents are returned.|
| **visBBoxDrawingCoords**|&;H2000|Returns numbers in the drawing coordinate system of the page or master whose shapes are being considered. By default, the returned numbers are drawing units in the local coordinate system of the parent of the considered shapes.|
| **visBBoxNoNonPrint**|&;H4000|Ignores the extents of shapes that are nonprinting. A shape is nonprinting if the value of its NonPrinting cell is non-zero or it belongs only to nonprinting layers.|
The extents rectangle is determined using the center of the shape's strokes; it does not take into account the width of the strokes. Nor does the rectangle include any area covered by shadows or line end markers. Microsoft Visio does not expose a means to determine a shape's "black bits" box, that is, the extents box adjusted to account for stroke widths, shadows, and line ends.

A shape may have control points or connection points that lie outside any of the bounding rectangles reported by the shape. You can determine the position of control points and connection points by querying results of the shape's cells.


## Example

The following procedure prints the dimensions of the bounding box of the selected shape in the Immediate window. If more than one shape is selected in the active window, a message box indicating an error is displayed. In all cases, results are reported in the drawing units of the page or master to which the shape belongs. This means that if the shape is a subshape of a group,  **visBBoxDrawingCoords** is passed as a flag to the **BoundingBox** method.

If the shape is a guide, the procedure passes  **visBBoxIncludeGuides** to the **BoundingBox** method so that the shape will be considered to have extent. Three rectangles are reported for the shape:




-  **visBBoxUprightWH** : an upright box that encloses the shape's width-height box
    
-  **visBBoxUprightText** : an upright box that encloses the shape's text box
    
-  **visBBoxExtents** : an upright box that encloses the shape's paths
    


To run this macro, make sure exactly one shape is selected on the Visio drawing page.




```vb
 
Public Sub BoundingBox_Example() 
 
 Dim vsoSelection As Visio.Selection 
 Set vsoSelection = ActiveWindow.Selection 
 vsoSelection.IterationMode = visSelModeSkipSub 
 
 If vsoSelection.Count <> 1 Then 
 MsgBox "BoundingBox_Example() expects exactly one selected shape." 
 
 Else 
 
 Dim vsoShape As Visio.Shape 
 Set vsoShape = vsoSelection(1) 
 Dim intFlags As Integer 
 intFlags = 0 
 
 If vsoShape.ContainingShape.Type = visTypeGroup Then 
 
 intFlags = visBBoxDrawingCoords 
 
 End If 
 
 If vsoShape.Type = visTypeGuide Then 
 
 intFlags = intFlags + visBBoxIncludeGuides 
 
 End If 
 
 Dim dblTop As Double 
 Dim dblBottom As Double 
 Dim dblLeft As Double 
 Dim dblRight As Double 
 
 vsoShape.BoundingBox intFlags + visBBoxUprightWH, dblLeft, dblBottom, dblRight, dblTop 
 Debug.Print "Upright WH "; _ 
 "dblLeft:" &; Application.FormatResult(dblLeft, "in", "", "#0.00 u"); _ 
 "dblBottom:" &; Application.FormatResult(dblBottom, "in", "", "#0.00 u"); _ 
 "dblRight:" &; Application.FormatResult(dblRight, "in", "", "#0.00 u"); _ 
 "dblTop:" &; Application.FormatResult(dblTop, "in", "", "#0.00 u") 
 
 vsoShape.BoundingBox intFlags + visBBoxUprightText, dblLeft, dblBottom, dblRight, dblTop 
 Debug.Print "Upright text "; _ 
 "dblLeft:" &; Application.FormatResult(dblLeft, "in", "", "#0.00 u"); _ 
 "dblBottom:" &; Application.FormatResult(dblBottom, "in", "", "#0.00 u"); _ 
 "dblRight:" &; Application.FormatResult(dblRight, "in", "", "#0.00 u"); _ 
 "dblTop:" &; Application.FormatResult(dblTop, "in", "", "#0.00 u") 
 
 vsoShape.BoundingBox intFlags + visBBoxExtents, dblLeft, dblBottom, dblRight, dblTop 
 Debug.Print "Bounding Box "; _ 
 "dblLeft:" &; Application.FormatResult(dblLeft, "in", "", "#0.00 u"); _ 
 "dblBottom:" &; Application.FormatResult(dblBottom, "in", "", "#0.00 u"); _ 
 "dblRight:" &; Application.FormatResult(dblRight, "in", "", "#0.00 u"); _ 
 "dblTop:" &; Application.FormatResult(dblTop, "in", "", "#0.00 u") 
 
 End If 
 
End Sub
```

The following macro uses the  **BoundingBox** method and the **ShapesOverlap()** function to determine if one shape (vsoShape2) overlaps another (vsoShape1).




```vb
Public Sub OverlappingShapes_Example() 
 
 Dim vsoShape1 As Visio.Shape 
 Dim vsoShape2 As Visio.Shape 
 Dim blsIsOverlapping As Boolean 
 
 
 Set vsoShape1 = Application.ActiveWindow.Page.Drop(Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Square"), 3, 9) 
 
 Set vsoShape2 = Application.ActiveWindow.Page.Drop(Application.Documents.Item("BASIC_U.VSS").Masters.ItemU("Pentagon"), 3, 8) 
 
 blsIsOverlapping = ShapesOverlap(vsoShape2, vsoShape1) 
 
 If blsIsOverlapping Then 
 Debug.Print "Shapes overlap." 
 Else 
 Debug.Print "Shapes do not overlap." 
 End If 
 
End Sub 
 
 
Private Function ShapesOverlap(vsoShape1 As IVShape, vsoShape2 As IVShape) As Boolean 
 
 Dim dblLeft1 As Double 
 Dim dblLeft2 As Double 
 Dim dblBottom1 As Double 
 Dim dblBottom2 As Double 
 Dim dblRight1 As Double 
 Dim dblRight2 As Double 
 Dim dblTop1 As Double 
 Dim dblTop2 As Double 
 
 vsoShape1.BoundingBox Flags + visBBoxExtents, dblLeft1, dblBottom1, dblRight1, dblTop1 
 vsoShape2.BoundingBox Flags + visBBoxExtents, dblLeft2, dblBottom2, dblRight2, dblTop2 
 
 If ((dblLeft2 >= dblLeft1 And dblLeft2 <= dblRight1) Or _ 
 (dblRight2 >= dblLeft1 And dblRight2 <= dblRight1)) And _ 
 ((dblTop2 >= dblBottom1 And dblTop2 <= dblTop1) Or _ 
 (dblBottom2 >= dblBottom1 And dblBottom2 <= dblTop1)) Then 
 ShapesOverlap = True 
 Else 
 ShapesOverlap = False 
 End If 
 
End Function
```


