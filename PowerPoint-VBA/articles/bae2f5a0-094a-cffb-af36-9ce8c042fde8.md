
# ColorSchemes.Count Property (PowerPoint)

 **Last modified:** July 28, 2015

Returns the number of objects in the specified collection. Read-only.

## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **ColorSchemes** object.


### Return Value

Long


## Example

This example closes all windows except the active window.


```
With Application.Windows

    For i = 2 To .Count

        .Item(2).Close

    Next

End With
```


## See also


#### Concepts


 [ColorSchemes Object](9b062448-88f5-b38d-2c76-330c691c9d72.md)
#### Other resources


 [ColorSchemes Object Members](df8e06a1-6c6b-1852-cb1f-e26929ba9bfa.md)
