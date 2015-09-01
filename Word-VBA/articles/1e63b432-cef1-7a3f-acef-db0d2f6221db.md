
# Pane.VerticalPercentScrolled Property (Word)

 **Last modified:** July 28, 2015

Returns or sets the vertical scroll position as a percentage of the document length. Read/write  **Long**.

## Syntax

 _expression_. **VerticalPercentScrolled**

 _expression_Required. A variable that represents a  ** [Pane](4a0c2690-d9d2-4e34-fef4-cc41365f5251.md)** object.


## Example

This example vertically scrolls the active pane of the window for Document1 to the end.


```
With Windows("Document1") 
 .Activate 
 .ActivePane.VerticalPercentScrolled = 100 
End With
```


## See also


#### Concepts


 [Pane Object](4a0c2690-d9d2-4e34-fef4-cc41365f5251.md)
#### Other resources


 [Pane Object Members](e0739460-3209-f981-71ea-80a5ea7f8935.md)
