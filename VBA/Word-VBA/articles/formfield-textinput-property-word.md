---
title: FormField.TextInput Property (Word)
keywords: vbawd10.chm153616395
f1_keywords:
- vbawd10.chm153616395
ms.prod: word
api_name:
- Word.FormField.TextInput
ms.assetid: 9a547325-344a-96ca-d22c-72c466d2522f
ms.date: 06/08/2017
---


# FormField.TextInput Property (Word)

Returns a  **[TextInput](textinput-object-word.md)** object that represents a text form field.


## Syntax

 _expression_ . **TextInput**

 _expression_ An expression that returns a **[FormField](formfield-object-word.md)** object.


## Remarks

If the  **TextInput** property is applied to a **FormField** object that isn't a drop-down form field, the property won't fail, but the **Valid** property for the returned object will be **False** .

Use the  **Result** property with the **FormField** object to return or set the contents of a **TextInput** object, as follows:


## Example

This example protects the active document for forms and deletes the contents of the form field named "Text1."


```vb
ActiveDocument.Protect Type:=wdAllowOnlyFormFields 
ActiveDocument.FormFields("Text1").TextInput.Clear
```

If the first form field in the active document is a text form field that accepts regular text, this example sets the contents of the form field.




```vb
Set myField = ActiveDocument.FormFields(1) 
If myField.Type = wdFieldFormTextInput And _ 
 myField.TextInput.Type = wdRegularText Then 
 myField.Result = "Hello" 
End If
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

