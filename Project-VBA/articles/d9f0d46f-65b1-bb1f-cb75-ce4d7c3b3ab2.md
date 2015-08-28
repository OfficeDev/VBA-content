
# ShapeRange.ZOrderPosition Property (Project)
Gets the position of the shape range in the z-order. Read-only  **Long**.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Property value](#sectionSection2)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ZOrderPosition**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks
<a name="sectionSection1"> </a>

To set the shape position in the z-order, use the  [ZOrder](e8badff9-fbe5-b6b8-8c33-68cfde3bef38.md) method.

The position of a shape in the z-order corresponds to the index number of the shape in the  **Shapes** collection. For example, if there are four shapes in the `myReport` report object, the expression `myReport.Shapes(1)` returns the shape at the back of the z-order, and the expression `myReport.Shapes(4)` returns the shape at the front of the z-order.

When you add a shape to a  **Shapes** collection, the shape is added to the front of the z-order by default.


## Property value
<a name="sectionSection2"> </a>

 **INT**


## See also
<a name="sectionSection2"> </a>


#### Other resources


 [ShapeRange Object](315031aa-4b8c-424b-26e7-ce15897beb05.md)
