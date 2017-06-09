---
title: Subdocuments Object (Word)
keywords: vbawd10.chm2440
f1_keywords:
- vbawd10.chm2440
ms.prod: word
ms.assetid: 8e14fced-19b3-2794-39a3-9e5ec15dd0ad
ms.date: 06/08/2017
---


# Subdocuments Object (Word)

A collection of  **[Subdocument](subdocument-object-word.md)** objects that represent the subdocuments in a range or document.


## Remarks

Use the  **Subdocuments** property to return the **Subdocuments** collection. The following example expands all the subdocuments in the active document.


```vb
ActiveDocument.Subdocuments.Expanded = True
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

Use  **Subdocuments** (Index), where Index is the index number, to return a single **Subdocument** object. The following example displays the path and file name of the first subdocument in the active document.




```vb
If ActiveDocument.Subdocuments(1).HasFile = True Then 
 MsgBox ActiveDocument.Subdocuments(1).Path &; _ 
 Application.PathSeparator _ 
 &; ActiveDocument.Subdocuments(1).Name 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

