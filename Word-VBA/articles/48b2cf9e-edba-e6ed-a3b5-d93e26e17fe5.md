
# Range.EndnoteOptions Property (Word)

 **Last modified:** July 28, 2015

Returns an  **EndnoteOptions** object that represents the endnotes in a range.

## Syntax

 _expression_. **EndnoteOptions**

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


## Example

This example sets the starting number for endnotes in section two of the active document to one if the starting number is not one.


```
Sub SetEndnoteOptionsRange() 
 With ActiveDocument.Sections(2).Range.EndnoteOptions 
 If .StartingNumber <> 1 Then 
 .StartingNumber = 1 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
