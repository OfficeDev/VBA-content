---
title: Requested Member of the Collection Does Not Exist
ms.prod: word
ms.assetid: 0053e3e4-8e33-c994-a910-572370dbbfb2
ms.date: 06/08/2017
---


# Requested Member of the Collection Does Not Exist

The "requested member of the collection does not exist" error occurs when you try to access an object that does not exist. For example, the following instruction may post an error if the active document does not contain at least one table.


```vb
Sub SelectTable() 
 ActiveDocument.Tables(1).Select 
End Sub
```


To avoid this error when accessing a member of a collection, ensure that the member exists prior to accessing the collection member. If you are accessing the member by index number, you can use the  **Count**property of the collection to determine if the member exists. The following example selects the first table, if there is at least one table in the active document.




```vb
Sub SelectFirstTable() 
 If ActiveDocument.Tables.Count > 0 Then 
 ActiveDocument.Tables(1).Select 
 Else 
 MsgBox "Document doesn't contain a table" 
 End If 
End Sub
```

If you are accessing a collection member by name, you can loop on the elements in a collection using a  **For Each...Next** loop to determine if the named member is part of the collection. For example, the following deletes the AutoCorrect entry named "acheive" if it is part of the **[AutoCorrectEntries](autocorrectentries-object-word.md)** collection. For more information, see  [Looping Through a Collection](looping-through-a-collection.md).



```vb
Sub DeleteAutoTextEntry() 
 Dim aceEntry As AutoCorrectEntry 
 For Each aceEntry In AutoCorrect.Entries 
 If aceEntry.Name = "acheive" Then aceEntry.Delete 
 Next aceEntry 
End Sub
```


