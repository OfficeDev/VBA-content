
# Shape.Line Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  **LineFormat** object that contains line formatting properties for the specified shape. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Line**

 _expression_A variable that represents a  ** [Shape](604029ce-9b2f-9748-5d4e-b458796fa2f0.md)** object.


## Remarks
<a name="sectionSection1"> </a>

For a line, the  **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.


## Example
<a name="sectionSection2"> </a>

This example adds a blue dashed line to  _myDocument_.


```
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

This example adds a cross to  _myDocument_ and then sets its border to be 8 points thick and red.




```
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Shape Object](604029ce-9b2f-9748-5d4e-b458796fa2f0.md)
#### Other resources


 [Shape Object Members](4aa8e2f4-5629-3922-11e4-df028bd1e1de.md)
