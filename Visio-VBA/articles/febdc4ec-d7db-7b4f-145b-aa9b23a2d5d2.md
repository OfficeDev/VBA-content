
# Selection.PrimaryItem Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns the  **Shape** object that is a **Selection** object's primary item. Read-only.


## Syntax

 _expression_. **PrimaryItem**

 _expression_A variable that represents a  **Selection** object.


### Return Value

Shape


## Remarks

In a drawing window, the primary selected item is shown with green selection handles and non-primary selected items are shown with blue selection handles. The outcome of some operations is affected by which selected item is the primary item. For example, the  **Align Shapes** command aligns non-primary selected items with the primary selected item.

If a  **Selection** object contains no **Shape** objects, or the primary **Shape** object is one that isn't enumerated given the **Selection** object's **IterationMode** property, the **PrimaryItem** property returns **Nothing** and raises no exception.

