
# Shape.ID Property (Publisher)

 **Last modified:** July 28, 2015

Returns a  **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.

## Syntax

 _expression_. **ID**

 _expression_A variable that represents a  **Shape** object.


## Example

This example displays the type for each shape on the first page of the active publication.


```
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```

