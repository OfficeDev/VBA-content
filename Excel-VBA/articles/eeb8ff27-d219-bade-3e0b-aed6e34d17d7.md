
# Application.Width Property (Excel)

Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the application window to its right edge.


## Syntax

 _expression_ . **Width**

 _expression_ A variable that represents an **Application** object.


## Remarks

 If the window is minimized, **Width** is read-only and returns the width of the window icon.


## Example

This example expands the active window to the maximum size available (assuming that the window isn't maximized).


```vb
With ActiveWindow 
 .WindowState = xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight 
 .Width = Application.UsableWidth 
End With
```


## See also


#### Concepts


[Application Object](19b73597-5cf9-4f56-8227-b5211f657f6f.md)
#### Other resources


[Application Object Members](4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e.md)
