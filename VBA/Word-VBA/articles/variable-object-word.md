---
title: Variable Object (Word)
keywords: vbawd10.chm2406
f1_keywords:
- vbawd10.chm2406
ms.prod: word
api_name:
- Word.Variable
ms.assetid: e6a75f54-6f91-75b4-7ca0-9be302e8dbe0
ms.date: 06/08/2017
---


# Variable Object (Word)

Represents a variable stored as part of a document. Document variables are used to preserve macro settings in between macro sessions. The  **Variable** object is a member of the **[Variables](variables-object-word.md)** collection. The **Variables** collection includes all the document variables in a document or template.


## Remarks

Use  **Variables** (Index), where Index is the document variable name or the index number, to return a single **Variable** object. The following example displays the value of the Temp document variable in the active document.


```vb
MsgBox ActiveDocument.Variables("Temp").Value
```

The index number represents the position of the document variable in the  **Variables** collection. The last variable added to the **Variables** collection is index number 1; the second-to-last variable added to the collection is index number 2, and so on. The following example displays the name of the first document variable in the active document.




```vb
MsgBox ActiveDocument.Variables(1).Name
```

Use the  **Add** method of the **Variables** collection to add a variable to a document. The following example adds a document variable named "Temp" with a value of 12 to the active document.




```vb
ActiveDocument.Variables.Add Name:="Temp", Value:="12"
```

If you try to add a document variable with a name that already exists in the  **Variables** collection, an error occurs. To avoid this error, you can enumerate the collection before adding any new variables. If the Blue document variable already exists in the active document, the following example sets its value to 6. If this variable does not already exist, this example adds it to the document and sets it to 6.




```vb
For Each aVar In ActiveDocument.Variables 
 If aVar.Name = "Blue" Then num = aVar.Index 
Next aVar 
If num = 0 Then 
 ActiveDocument.Variables.Add Name:="Blue", Value:=6 
Else 
 ActiveDocument.Variables(num).Value = 6 
End If
```

Document variables are invisible to the user unless a DOCVARIABLE field is inserted with the appropriate variable name. The following example adds a document variable named "Temp" to the active document and then inserts a DOCVARIABLE field to display the value in the variable.




```vb
With ActiveDocument 
 .Variables.Add Name:="Temp", Value:="12" 
 .Fields.Add Range:=Selection.Range, _ 
 Type:=wdFieldDocVariable, Text:="Temp" 
End With 
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
```

To add a document variable to a template, open the template as a document by using the  **OpenAsDocument** method. The following example stores the user name (from the **Options** dialog box) in the template attached to the active document.




```vb
ScreenUpdating = False 
With ActiveDocument.AttachedTemplate.OpenAsDocument 
 .Variables.Add Name:="UserName", Value:=Application.UserName 
 .Close SaveChanges:=wdSaveChanges 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

