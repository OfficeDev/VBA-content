---
title: FormField.CalculateOnExit Property (Word)
keywords: vbawd10.chm153616400
f1_keywords:
- vbawd10.chm153616400
ms.prod: word
api_name:
- Word.FormField.CalculateOnExit
ms.assetid: d92a165b-3138-9aae-bb98-08b7b01e52f8
ms.date: 06/08/2017
---


# FormField.CalculateOnExit Property (Word)

 **True** if references to the specified form field are automatically updated whenever the field is exited. Read/write **Boolean** .


## Syntax

 _expression_ . **CalculateOnExit**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

A REF field can be used to reference the contents of a form field. For example, {REF SubTotal} references the form field marked by the SubTotal bookmark.


## Example

This example keeps references to form fields in Form.doc from being automatically updated whenever the form field is exited.


```vb
Dim ffLoop As FormField 
 
For Each ffLoop In Documents("Form.doc").FormFields 
 ffLoop.CalculateOnExit = False 
Next ffLoop
```

This example adds a text form field and a REF field in a new document. Whenever text is typed and the Text1 field is exited, the REF field is automatically updated.




```vb
With Documents.Add 
 .FormFields.Add Range:=Selection.Range, _ 
 Type:=wdFieldFormTextInput 
 .Fields.Add Range:=Selection.Range, _ 
 Type:=wdFieldRef, Text:="Text1" 
 .FormFields("Text1").CalculateOnExit = True 
 .Protect Type:=wdAllowOnlyFormFields 
End With
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

