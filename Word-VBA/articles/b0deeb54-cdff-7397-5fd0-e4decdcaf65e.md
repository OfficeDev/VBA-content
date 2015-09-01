
# FormField.DropDown Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [DropDown](55233d61-d6d0-30f9-6825-ebbdbeb928b6.md)**object that represents a drop-down form field. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **DropDown**

 _expression_A variable that represents a  ** [FormField](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)** object.


## Remarks
<a name="sectionSection1"> </a>

If the  **DropDown** property is applied to a **FormField** object that isn't a drop-down form field, the property won't fail, but the **Valid**property for the returned object will be  **False**.


## Example
<a name="sectionSection2"> </a>

This example displays the text of the item selected in the drop-down form field named "Colors."


```
Dim ffDrop As FormField 
 
Set ffDrop = ActiveDocument.FormFields("Colors").DropDown 
 
MsgBox ffDrop.ListEntries(ffDrop.Value).Name
```

This example adds "Seattle" to the drop-down form field named "Places" in Form.doc.




```
With Documents("Form.doc").FormFields("Places") _ 
 .DropDown.ListEntries 
 .Add Name:="Seattle" 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [FormField Object](c3c07344-06b2-fe86-6fcb-b9c63a991bcc.md)
#### Other resources


 [FormField Object Members](e7d1b5d7-e1b3-b602-98c4-d0d4dc2288e5.md)
