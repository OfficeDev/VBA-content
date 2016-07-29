
# ContentControl.LockContents Property (Word)

Returns or sets a  **Boolean** that represents whether the user can edit the contents of a content control. Read/write.


## Syntax

 _expression_ . **LockContents**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The default value of this property is  **False** . This property corresponds to the **Contents cannot be edited** check box in the **Content Control Properties** dialog box.


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
