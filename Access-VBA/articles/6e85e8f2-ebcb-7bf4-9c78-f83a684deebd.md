
# ToggleButton.LabelY Property (Access)

 **Last modified:** July 28, 2015

The  **LabelY** property (along with the **LabelX** property) specifies the placement of the label for a new control. Read/write **Integer**.




## Syntax

 _expression_. **LabelY**

 _expression_A variable that represents a  **ToggleButton** object.


## Remarks

If the orientation is left to right for a form or report,  **LabelX** and **LabelY** behavior matches standard Microsoft Access left-to-right orientation. For more information about orientation, see the **Orientation** property.

If orientation is right to left, the origin of the coordinate system for  **LabelX** and **LabelY** is the upper right corner of the attached control. A negative number for **LabelX** places the label to the right of the control. A negative number for **LabelY** places the label above the control.

For General and Right alignment when orientation is RTL,  **LabelX** and **LabelY** specify the location of the upper-right corner of the label relative to the upper-right corner of the label's attached control. For Left and Center alignment, **LabelX** and **LabelY** specify the location of the upper-left corner and top center, respectively, of the label relative to the upper-right corner of the label's attached control.


## See also


#### Concepts


 [ToggleButton Object](1c20d809-d7db-e096-4328-ebb4d79e770e.md)
#### Other resources


 [ToggleButton Object Members](487101e7-c090-eb79-3671-5c9ce86cb6b0.md)
