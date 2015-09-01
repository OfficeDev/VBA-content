
# Window.Document Property (Word)

 **Last modified:** July 28, 2015

Returns a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object associated with the specified pane, window, or selection. Read-only.

## Syntax

 _expression_. **Document**

 _expression_A variable that represents a  ** [Window](d92f83f9-ae44-56c0-4584-7a9359253c6d.md)** object.


## Example

This example sets myDoc to the document associated with the active window. The focus is changed to the next window, and the window is split. The  **Activate** method is used to switch back to the original document.


```
Set myDoc = Application.ActiveWindow.Document 
If Windows.Count >= 2 Then 
 Application.ActiveWindow.Next.Activate 
 Application.ActiveWindow.Split = True 
 myDoc.Activate 
End If
```


## See also


#### Concepts


 [Window Object](d92f83f9-ae44-56c0-4584-7a9359253c6d.md)
#### Other resources


 [Window Object Members](c0dec747-3695-4f96-ea25-05b6494aad7e.md)
