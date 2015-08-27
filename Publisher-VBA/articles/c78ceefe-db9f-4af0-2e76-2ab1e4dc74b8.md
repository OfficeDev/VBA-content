
# ShapeNodes.Insert Method (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Inserts a new segment after the specified node of the freeform drawing.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Insert**( **_Index_**,  **_SegmentType_**,  **_EditingType_**,  **_X1_**,  **_Y1_**,  **_X2_**,  **_Y2_**,  **_X3_**,  **_Y3_**)

 _expression_A variable that represents a  **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the node after which the new node is to be inserted.|
|SegmentType|Required| **MsoSegmentType**|The type of segment to be added.|
|EditingType|Required| **MsoEditingType**|Specifies the editing type of the new node.|
|X1|Required| **Variant**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new node is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the first control point for the new segment.|
|Y1|Required| **Variant**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new node is  **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the first control point for the new segment.|
|X2|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the second control point for the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|Y2|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the second control point for the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|X3|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|
|Y3|Optional| **Variant**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment. If the EditingType of the new segment is  **msoEditingAuto**, do not specify a value for this argument.|

## Remarks
<a name="sectionSection1"> </a>

For the X1, Y1, X2, Y2, X3, and Y3 arguments, numeric values are evaluated in points; strings can be in any units supported by Publisher (for example, "2.5 in"). 

SegmentType can be one of these  **MsoSegmentType** constants.



| **msoSegmentCurve**|
| **msoSegmentLine**|
EditingType can be one of these  **MsoEditingType** constants.



| **msoEditingAuto**|Adds a node type appropriate to the segments being connected.|
| **msoEditingCorner**|Adds a corner node.|

## Example
<a name="sectionSection2"> </a>

This example adds a smooth node with a curved segment after node four in the third shape in the active publication. The shape must be a freeform drawing with at least four nodes.


```
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 .Insert Index:=4, _ 
 SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, _ 
 X1:=210, Y1:=100 
End With 

```

