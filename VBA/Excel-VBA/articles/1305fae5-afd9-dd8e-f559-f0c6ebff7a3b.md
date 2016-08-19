
# AxisTitle.Text Property (Excel)

Returns or sets the text for the specified object. Read/write  **String** .


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents an **AxisTitle** object.


## Example

This example sets the axis title text for the category axis in Chart1.


```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
End With
```


## See also


#### Concepts


[AxisTitle Object](563d3ba5-aa77-b6fc-236a-7838d75eaa53.md)
