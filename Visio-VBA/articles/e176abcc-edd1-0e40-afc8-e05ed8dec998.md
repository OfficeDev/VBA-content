
# Selection.Join Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Creates a new shape by joining selected shapes.


## Syntax

 _expression_. **Join**

 _expression_A variable that represents a  **Selection** object.


### Return Value

Nothing


## Remarks

Calling the  **Join** method is equivalent to clicking **Join** in the Microsoft Visio user interface (click **Operations** in the **Shape Design** group on the [Developer](1bdc55f5-8fc7-7257-03d5-c049eceb29ff.md) tab). The new shape inherits the text and formatting of the first selected shape and is the topmost shape in its containerâ€”thenth shape in the  **Shapes** collection of its containing shape, wheren = Count.

The original shapes are deleted and no shapes are selected when the operation is complete.

The  **Join** method and the **Combine** method are similar but differ in the following ways:




-  **Join** coalesces abutting line and curve segments in the original shapes into a single Geometry section in the resulting shape.
    
-  **Combine** produces a shape that has one Geometry section for each original shape. The resulting shape has holes in regions where the original shapes overlapped.
    


You might want to join shapes after importing a non-Visio drawing in which apparent polylines are represented by many independent shapes, each possessing a single line or curve segment. By joining the shapes that constitute a polyline in such a drawing, you can replace many single-segment shapes with one multiple-segment shape.

