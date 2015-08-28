
# Master.ColorScheme Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns or sets the  ** [ColorScheme](c1945542-b628-e2b1-5114-e064f0563a01.md)**object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.

## Syntax

 _expression_. **ColorScheme**

 _expression_A variable that represents a  **Master** object.


### Return Value

ColorScheme


## Example

This example sets the title color to green for slides one and three in the active presentation.


```
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```


## See also


#### Concepts


 [Master Object](22e8805e-6469-1a34-7f7b-f1ea5c6c49ff.md)
#### Other resources


 [Master Object Members](156762f4-61b8-43d0-2ce3-3069184cc225.md)
