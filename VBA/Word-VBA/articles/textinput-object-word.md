---
title: TextInput Object (Word)
keywords: vbawd10.chm2343
f1_keywords:
- vbawd10.chm2343
ms.prod: word
api_name:
- Word.TextInput
ms.assetid: d7f6531a-4da2-ccc4-29b3-ad79ca7b18de
ms.date: 06/08/2017
---


# TextInput Object (Word)

Represents a single text form field.


## Remarks

Use  **FormFields** (Index), where Index is either the bookmark name associated with the text form field or the index number, to return a **FormField** object. Use the **TextInput** property with the **FormField** object to return a **TextInput** object. The following example deletes the contents of the text form field named "Text1" in the active document.


```vb
ActiveDocument.FormFields("Text1").TextInput.Clear
```

The index number represents the position of the form field in the  **FormFields** collection. The following example checks the type of the first form field in the active document. If the form field is a text form field, the example sets "Mission Critical" as the value of the field.




```vb
If ActiveDocument.FormFields(1).Type = wdFieldFormTextInput Then 
 ActiveDocument.FormFields(1).Result = "Mission Critical" 
End If
```

The following example determines whether the  _ffield_ variable represents a valid text form field in the active document before it sets the default text.




```vb
Set ffield = ActiveDocument.FormFields(1).TextInput 
If ffield.Valid = True Then 
 ffield.Default = "Type your name here" 
Else 
 MsgBox "First field is not a text box" 
End If
```

Use the  **Add** method with the **[FormFields](formfields-object-word.md)** object to add a text form field. The following example adds a text form field at the beginning of the active document and then sets the name of the form field to "FirstName."




```vb
Set ffield = ActiveDocument.FormFields.Add( _ 
 Range:=ActiveDocument.Range(Start:=0, End:=0), _ 
 Type:=wdFieldFormTextInput) 
ffield.Name = "FirstName"
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


