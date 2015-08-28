
# Shape.ZOrderPosition Property (Project)
Gets the position of the shape in the z-order. Read-only  **Long**.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Property value](#sectionSection2)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ZOrderPosition**

 _expression_A variable that represents a  **Shape** object.


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


 [Shape Object](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)
 [Shapes Object](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)
