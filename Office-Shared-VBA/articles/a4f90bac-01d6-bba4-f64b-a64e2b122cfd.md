
# CustomXMLPart Object (Office)

 **Last modified:** July 28, 2015

Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.

## Example

The following example adds a part to a  **CustomXMLPart** object.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## See also


#### Concepts


 [Object Model Reference](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Other resources


 [CustomXMLPart Object Members](76fe85f4-5a35-7d12-2989-6f17a094dcdf.md)
