
# Axis.CategoryType Property (Word)

Returns or sets the category axis type. Read/write  **[XlCategoryType](10dad161-2a90-7915-51bb-ddc69427c003.md)** .


## Syntax

 _expression_ . **CategoryType**

 _expression_ A variable that represents an **[Axis](3a7ad7d8-d397-a79a-eb6a-a5f0822cbe5d.md)** object.


## Remarks

You cannot set this property for a value axis.


## Example

The following example sets the category axis for the first chart in the active document to use a time scale, using months as the base unit.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Axis Object](3a7ad7d8-d397-a79a-eb6a-a5f0822cbe5d.md)
#### Other resources


[Axis Object Members](44fa1b67-2a56-3d92-cb63-4144e1bb7282.md)
