
# Axes.Parent Property (Excel)

 **Last modified:** July 28, 2015

Returns the parent object for the specified object. Read-only.

## Syntax

 _expression_. **Parent**

 _expression_A variable that represents an  **Axes** object.


## Example

This example displays the name of the chart that contains  `myAxis`.


```
Sub DisplayParentName() 
 
 Set myAxis = Charts(1).Axes(xlValue) 
 MsgBox myAxis.Parent.Name 
 
End Sub
```


## See also


#### Concepts


 [Axes Collection](581e51e5-3dbb-7f0c-a87d-2d44f67dad0b.md)
#### Other resources


 [Axes Object Members](10a6fffe-65ff-e9b2-813c-357664e276a5.md)
