
# TextFrame2.Ruler Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns a  **Ruler2** object that represents the ruler for the specified text. Read-only.

## Syntax

 _expression_. **Ruler**

 _expression_An expression that returns a  **TextFrame2** object.


### Return Value

Ruler2


## Example

This example shows how to set a left-aligned tab stop at 2 inches (144 points) for the text in shape one on slide one in the active presentation.


```
Public Sub Ruler_Example() 
 
    Dim pptSlide As Slide 
    Set pptSlide = ActivePresentation.Slides(1) 
    pptSlide.Shapes(1).TextFrame2.Ruler.TabStops.Add ppTabStopLeft, 144 
 
End Sub
```


## See also


#### Concepts


 [TextFrame2 Object](ae017598-8330-4673-db1a-53b284acb709.md)
#### Other resources


 [TextFrame2 Object Members](bce672a4-b108-b223-7e65-71f07d7f4197.md)
