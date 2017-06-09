---
title: FormField.DropDown Property (Word)
keywords: vbawd10.chm153616397
f1_keywords:
- vbawd10.chm153616397
ms.prod: word
api_name:
- Word.FormField.DropDown
ms.assetid: b0deeb54-cdff-7397-5fd0-e4decdcaf65e
ms.date: 06/08/2017
---


# FormField.DropDown Property (Word)

Returns a  **[DropDown](dropdown-object-word.md)** object that represents a drop-down form field. Read-only.


## Syntax

 _expression_ . **DropDown**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


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


[FormField Object](formfield-object-word.md)

