
# FormField.DropDown Property (Word)

Returns a  **[DropDown](55233d61-d6d0-30f9-6825-ebbdbeb928b6.md)** object that represents a drop-down form field. Read-only.


## Syntax

 _expression_ . **DropDown**

 _expression_ A variable that represents a **[FormField](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)** object.


## Remarks

If the  **DropDown** property is applied to a **FormField** object that isn't a drop-down form field, the property won't fail, but the **Valid** property for the returned object will be **False** .


## Example

This example displays the text of the item selected in the drop-down form field named "Colors."


```vb
Dim ffDrop As FormField 
 
Set ffDrop = ActiveDocument.FormFields("Colors").DropDown 
 
MsgBox ffDrop.ListEntries(ffDrop.Value).Name
```

This example adds "Seattle" to the drop-down form field named "Places" in Form.doc.




```vb
With Documents("Form.doc").FormFields("Places") _ 
 .DropDown.ListEntries 
 .Add Name:="Seattle" 
End With
```


## See also


#### Concepts


[FormField Object](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)
#### Other resources


[FormField Object Members](e7d1b5d7-e1b3-b602-98c4-d0d4dc2288e5.md)
