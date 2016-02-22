
# ContentControl.Range Property (Word)

Returns a  **[Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** that represents the contents of the content control in the active document. Read-only.


## Syntax

 _expression_ . **Range**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

Use the  **Range** property to access the text, the formatting for the text, and other text properties. For more information, see[Working with Range Objects](http://msdn.microsoft.com/library/9e240aa7-8608-9d70-aee3-2e202687459e%28Office.15%29.aspx).


## Example

The following example inserts a date content control into the active document, and then sets the contents of the content control and specifies that the user cannot edit the contents or delete the control from the document.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls _ 
 .Add(wdContentControlDate) 
 
objCC.Range.Text = "January 1, 2007" 
objCC.LockContents = True 
objCC.LockContentControl = True
```


## See also


#### Concepts


[ContentControl Object](783dec26-9b63-11f8-6187-985f9c815f27.md)
#### Other resources


[ContentControl Object Members](d5aa195c-8d7a-0bad-09fa-6f1bfc9828cc.md)
