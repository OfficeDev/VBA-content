
# ShapeNodes.SetPosition Method (PowerPoint)

 **Last modified:** July 28, 2015

Sets the location of the node specified by  **Index**. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.

## Syntax

 _expression_. **SetPosition**( **_Index_**,  **_X1_**,  **_Y1_**)

 _expression_A variable that represents a  **ShapeNodes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The node whose position is to be set.|
| Y1|Required| **Single**|The x-position (in points) of the new node relative to the upper-left corner of the document.|
| Y1|Required| **Single**|The y-position (in points) of the new node relative to the upper-left corner of the document.|

## Example

This example moves node two in shape three on  `myDocument` to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    pointsArray = .Item(2).Points

    currXvalue = pointsArray(1, 1)

    currYvalue = pointsArray(1, 2)

    .SetPosition 2, currXvalue + 200, currYvalue + 300

End With
```


## See also


#### Concepts


 [ShapeNodes Object](493bacfe-eb8c-2064-46ec-c19e58e9b1ce.md)
#### Other resources


 [ShapeNodes Object Members](790cc468-e7eb-97f5-ac0a-5ecc526ebfd2.md)
