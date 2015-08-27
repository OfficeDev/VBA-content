
# Shapes.Paste Method (Publisher)

 **Last modified:** July 28, 2015

Pastes the shapes or text on the Clipboard into the specified  ** [Shapes](52e069a6-d54b-a11a-1cba-96174329cb02.md)** collection, at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a ** [ShapeRange](c85967c9-af43-747d-7e0b-64ddc22c84be.md)** object that represents the pasted objects.

## Syntax

 _expression_. **Paste**

 _expression_A variable that represents a  **Shapes** object.


### Return Value

ShapeRange


## Example

This example copies shape one on page one in the active publication to the Clipboard and then pastes it into page two.


```
With ActiveDocument 
 .Pages(1).Shapes(1).Copy 
 .Pages(2).Shapes.Paste 
End With
```

