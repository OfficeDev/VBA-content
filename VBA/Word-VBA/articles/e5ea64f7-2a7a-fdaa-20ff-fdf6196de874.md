
# Application.UsableWidth Property (Word)

Returns the maximum width (in points) to which you can set the width of a Microsoft Word document window. Read-only  **Long** .


## Syntax

 _expression_ . **UsableWidth**

 _expression_ A variable that represents an **[Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


## Example

This example sets the size of the active document window to one quarter of the maximum allowable screen area.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Top = 5 
 .Left = 5 
 .Height = (Application.UsableHeight*0.5) 
 .Width = (Application.UsableWidth*0.5) 
End With
```

This example displays the size of the working area in the active document window.




```vb
With ActiveDocument.ActiveWindow 
 MsgBox "Working area height = " _ 
 &; .UsableHeight &; vbLf _ 
 &; "Working area width = " _ 
 &; .UsableWidth 
End With
```


## See also


#### Concepts


[Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


[Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
