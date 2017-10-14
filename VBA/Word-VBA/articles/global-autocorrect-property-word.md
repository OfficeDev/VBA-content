---
title: Global.AutoCorrect Property (Word)
keywords: vbawd10.chm163119114
f1_keywords:
- vbawd10.chm163119114
ms.prod: word
api_name:
- Word.Global.AutoCorrect
ms.assetid: 3565507b-c2b7-da6c-a725-ab925d695c6d
ms.date: 06/08/2017
---


# Global.AutoCorrect Property (Word)

Returns an  **[AutoCorrect](autocorrect-object-word.md)** object that contains the current AutoCorrect options, entries, and exceptions. Read-only.


## Syntax

 _expression_ . **AutoCorrect**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Example

This example adds an AutoCorrect replacement entry. After this code runs, every instance of "sr" that's typed in a document will automatically be replaced with "Stella Richards."


```
AutoCorrect.Entries.Add Name:= "sr", Value:= "Stella Richards"
```

This example deletes the specified AutoCorrect entry it if it exists.




```vb
Dim strInput as String 
Dim aceLoop as AutoCorrectEntry 
Dim blnMatch as Boolean 
Dim intConfirm as Integer 
 
blnMatch = False 
 
strInput = InputBox("Enter the AutoCorrect entry to delete.") 
 
For Each aceLoop in AutoCorrect.Entries 
 With aceLoop 
 If .Name = strInput Then 
 blnMatch = True 
 intConfirm = _ 
 MsgBox("Are you sure you want to delete " &; _ 
 .Name, 4) 
 If intConfirm = vbYes Then 
 .Delete 
 End If 
 End If 
 End With 
Next aceLoop 
 
If blnMatch <> True Then 
 MsgBox "There was no AutoCorrect entry: " &; strInput 
End If
```


## See also


#### Concepts


[Global Object](global-object-word.md)

