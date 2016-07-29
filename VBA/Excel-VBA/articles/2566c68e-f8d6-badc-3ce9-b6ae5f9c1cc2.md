
# ShadowFormat Object (Excel)

Represents shadow formatting for a shape.


## Remarks

Use the  **[Shadow](e44d59d4-5e85-3c78-b3a4-eabac9f2b86f.md)** property to return a **ShadowFormat** object.


## Example

 The following example adds a shadowed rectangle to _myDocument_ . The semitransparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 50, 50, 100, 200).Shadow 
 .ForeColor.RGB = RGB(0, 0, 128) 
 .OffsetX = 5 
 .OffsetY = -3 
 .Transparency = 0.5 
 .Visible = True 
End With
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
[ShadowFormat Object Members](5512df5b-d899-7942-1309-4cf8d28fe96a.md)
