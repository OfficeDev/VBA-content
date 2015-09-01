
# Fields.Update Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Updates the result of the fields object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Update**

 _expression_Required. A variable that represents a  ** [Fields](c79065bb-ba29-22fd-a9d7-90bb10550035.md)** collection.


### Return Value

Long


## Remarks
<a name="sectionSection1"> </a>

Returns 0 (zero) if no errors occur when the fields are updated, or returns a  **Long** that represents the index of the first field that contains an error.


## Example
<a name="sectionSection2"> </a>

This example updates all the fields in the main story (that is, the main body) of the active document. A return value of 0 (zero) indicates that the fields were updated without error.


```
If ActiveDocument.Fields.Update = 0 Then 
 MsgBox "Update Successful" 
Else 
 MsgBox "Field " &amp; ActiveDocument.Fields.Update &amp; _ 
 " has an error" 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Fields Collection Object](c79065bb-ba29-22fd-a9d7-90bb10550035.md)
#### Other resources


 [Fields Object Members](b480b07e-2a71-0e3d-113c-962fcd484bd4.md)
