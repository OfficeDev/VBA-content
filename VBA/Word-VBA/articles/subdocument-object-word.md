---
title: Subdocument Object (Word)
keywords: vbawd10.chm2441
f1_keywords:
- vbawd10.chm2441
ms.prod: word
api_name:
- Word.Subdocument
ms.assetid: ed966369-34f6-ef0c-d6d6-4c86baff4793
ms.date: 06/08/2017
---


# Subdocument Object (Word)

Represents a subdocument within a document or range. The  **Subdocument** object is a member of the **[Subdocuments](subdocuments-object-word.md)** collection. The **Subdocuments** collection includes all the subdocuments in the a range or document.


## Remarks

Use  **Subdocuments** (Index), where Index is the index number, to return a single **Subdocument** object. The following example displays the path and file name of the first subdocument in the active document.


```vb
If ActiveDocument.Subdocuments(1).HasFile = True Then 
 MsgBox ActiveDocument.Subdocuments(1).Path &; _ 
 Application.PathSeparator &; _ 
 ActiveDocument.Subdocuments(1).Name 
End If
```

Use the  **AddFromFile** or **AddFromRange** method to add a subdocument to a document. The following example adds a subdocument named "Setup.doc" at the end of the active document.




```vb
ActiveDocument.Subdocuments.Expanded = True 
Selection.EndKey Unit:=wdStory 
Selection.InsertParagraphBefore 
ActiveDocument.Subdocuments.AddFromFile Name:="C:\Temp\Setup.doc"
```

The following example applies the Heading 1 style to the first paragraph in the selection and then creates a subdocument for the contents of the selection.




```vb
Selection.Paragraphs(1).Style = wdStyleHeading1 
With ActiveDocument.Subdocuments 
 .Expanded = True 
 .AddFromRange Range:=Selection.Range 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


