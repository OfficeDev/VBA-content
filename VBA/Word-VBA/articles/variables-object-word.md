---
title: Variables Object (Word)
keywords: vbawd10.chm2405
f1_keywords:
- vbawd10.chm2405
ms.prod: word
ms.assetid: 9719d0d4-319d-c710-d243-12a9dee45880
ms.date: 06/08/2017
---


# Variables Object (Word)

A collection of  **[Variable](variable-object-word.md)** objects that represent the variables added to a document or template. Document variables are used to preserve macro settings in between macro sessions.


## Remarks

Use the  **Variables** property to return the **Variables** collection. The following example displays the number of variables in the document named "Sales.doc."


```vb
MsgBox Documents("Sales.doc").Variables.Count &; " variables"
```

Use the  **Add** method to add a variable to a document. The following example adds a document variable named "Temp" with a value of 12 to the active document.




```vb
ActiveDocument.Variables.Add Name:="Temp", Value:="12"
```

If you try to add a document variable with a name that already exists in the  **Variables** collection, an error occurs. To avoid this error, you can enumerate the collection before adding any new variables. If the Blue document variable already exists in the active document, the following example sets its value to 6. If this variable doesn't already exist, this example adds it to the document and sets it to 6.




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

Use  **Variables** (Index), where Index is the document variable name or the index number, to return a single **Variable** object. The following example displays the value of the Temp document variable in the active document.




```vb
MsgBox ActiveDocument.Variables("Temp").Value
```

The index number represents the position of the document variable in the  **Variables** collection. The first variable added to the **Variables** collection is index number 1; the second variable added to the collection is index number 2, and so on. The following example displays the name of the first document variable in the active document.




```vb
MsgBox ActiveDocument.Variables(1).Name
```

To add a variable to a template, open the template as a document by using the  **OpenAsDocument** method. The following example stores the user name (from the **Options** dialog box) in the template attached to the active document.




```vb
ScreenUpdating = False 
With ActiveDocument.AttachedTemplate.OpenAsDocument 
 .Variables.Add Name:="UserName", Value:= Application.UserName 
 .Close SaveChanges:=wdSaveChanges 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


