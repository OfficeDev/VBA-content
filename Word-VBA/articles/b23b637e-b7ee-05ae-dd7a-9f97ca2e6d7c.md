
# Browser.Previous Method (Word)

 **Last modified:** July 28, 2015

Moves the selection to the previous item indicated by the browser target. Use the  **Target** property to change the browser target.

## Syntax

 _expression_. **Previous**

 _expression_Required. A variable that represents a  ** [Browser](447bcab6-cfb2-77b0-9bbd-35e774417a60.md)** object.


## Example

This example moves the insertion point into the first cell (the cell in the upper-left corner) of the previous table.


```
With Application.Browser 
 .Target = wdBrowseTable 
 .Previous 
End With
```


## See also


#### Concepts


 [Browser Object](447bcab6-cfb2-77b0-9bbd-35e774417a60.md)
#### Other resources


 [Browser Object Members](ab97f30f-71c5-4360-0f6d-c47b7b45f0a3.md)
