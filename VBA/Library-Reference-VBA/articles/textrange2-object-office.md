---
title: TextRange2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.TextRange2
ms.assetid: a6a59c9b-9b64-c1e2-2e98-a1f99025c877
---


# TextRange2 Object (Office)

Represents the text frame in a  **Shape** or **ShapeRange** objects.


## Remarks

This object contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame. Use the  **TextFrame2** property to return a **TextFrame2** object.


## Example

The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame. 


```
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeRectangle, _ 
 0, 0, 250, 140).TextFrame2 
 .TextRange.Text = "Here is some test text" 
 .MarginBottom = 10 
 .MarginLeft = 10 
 .MarginRight = 10 
 .MarginTop = 10 
End With 

```


## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[TextRange2 Object Members](http://msdn.microsoft.com/library/26daffff-b9ef-fd94-f5b7-ed3a09840cb6%28Office.15%29.aspx)
