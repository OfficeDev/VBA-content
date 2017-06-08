---
title: Looping Through a Collection
ms.prod: word
ms.assetid: 68a4644f-888a-d46c-3c84-8a11f5993ec6
ms.date: 06/08/2017
---


# Looping Through a Collection

There are several different ways you can loop on the elements of a collection. However, the recommended method for looping on a collection is to use the  **For Each...Next** loop. In this structure, Visual Basic repeats a block of statements for each object in a collection. The following example displays the name of each document in the **[Documents](documents-object-word.md)** collection.


```vb
Sub LoopThroughOpenDocuments() 
 Dim docOpen As Document 
 
 For Each docOpen In Documents 
 MsgBox docOpen.Name 
 Next docOpen 
End Sub
```


Instead of displaying each element name in a message box, you can use an array to store the information. This example uses an array to store the name of each bookmark contained in the active document.




```vb
Sub LoopThroughBookmarks() 
 Dim bkMark As Bookmark 
 Dim strMarks() As String 
 Dim intCount As Integer 
 
 If ActiveDocument.Bookmarks.Count > 0 Then 
 ReDim strMarks(ActiveDocument.Bookmarks.Count - 1) 
 intCount = 0 
 For Each bkMark In ActiveDocument.Bookmarks 
 strMarks(intCount) = bkMark.Name 
 intCount = intCount + 1 
 Next bkMark 
 End If 
End Sub
```

You can loop through a collection to conditionally perform a task on members of the collection. For example, the following code updates the DATE fields in the active document.



```vb
Sub UpdateDateFields() 
 Dim fldDate As Field 
 
 For Each fldDate In ActiveDocument.Fields 
 If InStr(1, fldDate.Code, "Date", 1) Then fldDate.Update 
 Next fldDate 
End Sub
```

You can loop through a collection to determine if an element exists. For example, the following code displays a message if an AutoText entry named "Filename" is part of the  **[AutoTextEntries](autotextentries-object-word.md)** collection.



```vb
Sub FindAutoTextEntry() 
 Dim atxtEntry As AutoTextEntry 
 
 For Each atxtEntry In ActiveDocument.AttachedTemplate.AutoTextEntries 
 If atxtEntry.Name = "Filename" Then _ 
 MsgBox "The Filename AutoText entry exists." 
 Next atxtEntry 
End Sub
```


