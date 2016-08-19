
# Application.ActiveWindow Property (Publisher)

Returns a  **[Window](342d77cd-5556-6ac3-a828-b1b60380f910.md)** object that represents the window with the focus. Because Microsoft Publisher only has one window, there is only one **Window** object to return.


## Syntax

 _expression_. **ActiveWindow**

 _expression_A variable that represents an  **Application** object.


## Example

This example displays the active window's caption.


```vb
Sub CurrentCaption() 
 
 MsgBox ActiveDocument.ActiveWindow.Caption 
 
End Sub
```


## See also


#### Concepts


 [Application Object](acfc7efb-e6a5-a89a-3aee-3cb4af2f3508.md)
