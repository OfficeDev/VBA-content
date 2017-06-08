---
title: Returning Text from a Document
ms.prod: word
ms.assetid: bacf3de8-ae60-2f27-fa28-e53518e04be2
ms.date: 06/08/2017
---


# Returning Text from a Document

Use the  **Text**property to return text from a  **[Range](range-object-word.md)** object or **[Selection](selection-object-word.md)** object. The following example selects the next paragraph formatted with the Heading 1 style. The contents of the **Text** property are displayed by the **MsgBox** function.


```vb
Sub FindHeadingStyle() 
 With Selection.Find 
 .ClearFormatting 
 .Style = wdStyleHeading1 
 .Execute FindText:="", Format:=True, _ 
 Forward:=True, Wrap:=wdFindStop 
 If .Found = True Then MsgBox Selection.Text 
 End With 
End Sub
```


The following instruction returns and displays the selected text.




```vb
Sub ShowSelection() 
 Dim strText As String 
 strText = Selection.Text 
 MsgBox strText 
End Sub
```

The following example returns the first word in the active document. Each item in the  **[Words](words-object-word.md)** collection is a  **Range**object that represents one word.



```vb
Sub ShowFirstWord() 
 Dim strFirstWord As String 
 strFirstWord = ActiveDocument.Words(1).Text 
 MsgBox strFirstWord 
End Sub
```

The following example returns the text associated with the first bookmark in the active document.



```vb
Sub ShowFirstBookmark() 
 Dim strBookmark As String 
 If ActiveDocument.Bookmarks.Count > 0 Then 
 strBookmark = ActiveDocument.Bookmarks(1).Range.Text 
 MsgBox strBookmark 
 End If 
End Sub
```


