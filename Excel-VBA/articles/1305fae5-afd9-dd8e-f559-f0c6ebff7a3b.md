
# AxisTitle.Text Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets the text for the specified object. Read/write  **String**.

## Syntax

 _expression_. **Text**

 _expression_A variable that represents an  **AxisTitle** object.


## Example

This example sets the axis title text for the category axis in Chart1.


```
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
End With
```


## See also


#### Concepts


 [AxisTitle Object](563d3ba5-aa77-b6fc-236a-7838d75eaa53.md)
#### Other resources


 [AxisTitle Object Members](84970b5a-91a1-b785-5632-97a0de4410f2.md)
