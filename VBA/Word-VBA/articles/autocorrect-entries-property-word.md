---
title: AutoCorrect.Entries Property (Word)
keywords: vbawd10.chm155779078
f1_keywords:
- vbawd10.chm155779078
ms.prod: word
api_name:
- Word.AutoCorrect.Entries
ms.assetid: eaf66013-5417-742b-9bf1-cbf83626a8e5
ms.date: 06/08/2017
---


# AutoCorrect.Entries Property (Word)

Returns an  **[AutoCorrectEntries](autocorrectentries-object-word.md)** collection that represents the current list of AutoCorrect entries.


## Syntax

 _expression_ . **Entries**

 _expression_ A variable that represents an **[AutoCorrect](autocorrect-object-word.md)** object.


## Remarks

This list corresponds to the list of AutoCorrect entries on the  **AutoCorrect** tab in the **AutoCorrect** dialog box. Read-only. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the total number of AutoCorrect entries.


```vb
MsgBox AutoCorrect.Entries.Count
```

This example deletes the specified AutoCorrect entry if it exists.




```vb
Dim strEntry As String 
Dim acEntry As AutoCorrectEntry 
Dim blnMatch As Boolean 
Dim intResponse As Integer 
 
strEntry = InputBox("Enter the AutoCorrect entry to delete.") 
blnMatch = False 
 
For Each acEntry in AutoCorrect.Entries 
 If acEntry.Name = strEntry Then 
 blnMatch = True 
 intResponse = _ 
 MsgBox("Are you sure you want to delete " _ 
 &; acEntry.Name, 4) 
 If intResponse = vbYes Then 
 acEntry.Delete 
 End If 
 End If 
Next acEntry 
 
If blnMatch <> True Then 
 MsgBox "There was no AutoCorrect entry: " &; strEntry 
End If
```


## See also


#### Concepts


[AutoCorrect Object](autocorrect-object-word.md)

