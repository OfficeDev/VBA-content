
# SlideShowWindow.View Property (PowerPoint)

Returns a  **[SlideShowView](403b30ef-b12f-3a3c-e8d8-19189fd762fe.md)** object. Read-only.


## Syntax

 _expression_. **View**

 _expression_ A variable that represents a **SlideShowWindow** object.


### Return Value

SlideShowView


## Example

This example uses the  **View** property to exit the current slide show, sets the view in the active window to slide view, and then displays slide three.


```vb
Application.SlideShowWindows(1).View.Exit

With Application.ActiveWindow

    .ViewType = ppViewSlide

    .View.GotoSlide 3

End With
```


## See also


#### Concepts


[SlideShowWindow Object](22468489-d4a2-ffea-7479-53ecb8d5da29.md)
#### Other resources


[SlideShowWindow Object Members](7b2d0120-81a7-3232-fc38-f932f351523a.md)
