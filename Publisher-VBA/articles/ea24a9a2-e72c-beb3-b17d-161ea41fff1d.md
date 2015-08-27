
# Shapes.BuildFreeform Method (Publisher)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Builds a freeform object. Returns a  [FreeformBuilder](542df9f7-f636-a98e-01de-11005b5797cc.md)object that represents the freeform as it is being built.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BuildFreeform**( **_EditingType_**,  **_X1_**,  **_Y1_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|EditingType|Required| **MsoEditingType**|Specifies the editing type of the first node.|
|X1|Required| **Variant**|The horizontal position of the first node in the freeform drawing relative to the upper-left corner of the page.|
|Y1|Required| **Variant**|The vertical position of the first node in the freeform drawing relative to the upper-left corner of the page.|

### Return Value

FreeformBuilder


## Remarks
<a name="sectionSection1"> </a>

The EditingType parameter can be one of the following  **MsoEditingType** constants declared in the Microsoft Office type library.



| **msoEditingAuto**|Adds a node type appropriate to the segments being connected.|
| **msoEditingCorner**|Adds a corner node.|

## Example
<a name="sectionSection2"> </a>

For the X1 and Y1 arguments, numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").



Use the  ** [AddNodes](29906bde-e6a6-f661-0f3f-085f39653e42.md)**method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the  [ConvertToShape](1cb490af-40be-b03f-2f8d-04b1015fbde3.md)method to convert the  **FreeformBuilder** object into a **Shape** object that has the geometric description you've defined in the **FreeformBuilder** object.




```
' Add a new freeform object. 
With ActiveDocument.Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 
 

```

