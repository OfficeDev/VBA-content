
# ChartColorFormat.RGB Property (Word)

 **Last modified:** July 28, 2015

Returns the red-green-blue value of the specified color. Read-only  **Long**.

## Syntax

 _expression_. **RGB**

 _expression_A variable that represents a  ** [ChartColorFormat](8bc25b6c-3691-fc85-fcc6-d21ed3f903b9.md)** object.


## Example

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to the chart area foreground fill color, for the first chart group of the first chart in the active document.


```
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .ChartGroups(1).HasUpDownBars = True 
 .ChartGroups(1).DownBars.Interior.Pattern = xlPatternCrissCross 
 .ChartGroups(1).DownBars.Interior.PatternColor = _ 
 .ChartArea.Fill.ForeColor.RGB 
 End With 
 End If 
End With
```


## See also


#### Concepts


 [ChartColorFormat Object](8bc25b6c-3691-fc85-fcc6-d21ed3f903b9.md)
#### Other resources


 [ChartColorFormat Object Members](f3bbb759-bbc1-366c-a6ce-151c47580fa7.md)
