---
title: StyleSheet.Title Property (Word)
keywords: vbawd10.chm166658054
f1_keywords:
- vbawd10.chm166658054
ms.prod: word
api_name:
- Word.StyleSheet.Title
ms.assetid: 050e5915-2e92-7023-fb64-e122bfc4dd38
ms.date: 06/08/2017
---


# StyleSheet.Title Property (Word)

Returns or sets a  **String** representing the title of a Web style sheet. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ Required. A variable that represents a **[StyleSheet](stylesheet-object-word.md)** object.


## Example

This example assigns titles to the first three Web style sheets attached to the active document. This example assumes that there are three style sheets attached to the active document.


```vb
Sub AssignCSSTitle() 
 ActiveDocument.StyleSheets.Item(1).Title = "New Look Stylesheet" 
 ActiveDocument.StyleSheets.Item(2).Title = "Standard Web Stylesheet" 
 ActiveDocument.StyleSheets.Item(3).Title = "Definitions Stylesheets" 
End Sub
```

This example creates a list of Web style sheets attached to the active document and places the list in a new document. This example assumes there are one or more Web style sheets attached to the active document.




```vb
Sub CSSTitles() 
 Dim docNew As Document 
 Dim styCSS As StyleSheet 
 
 Set docNew = Documents.Add 
 
 With docNew.Range(Start:=0, End:=0) 
 .InsertAfter "CSS Name : Assigned to " &; ActiveDocument.Name _ 
 &; vbTab &; "Title" 
 .InsertParagraphAfter 
 For Each styCSS In ActiveDocument.StyleSheets 
 .InsertAfter styCSS.Name &; vbTab &; styCSS.Title 
 .InsertParagraphAfter 
 Next styCSS 
 .ConvertToTable 
 End With 
End Sub
```


## See also


#### Concepts


[StyleSheet Object](stylesheet-object-word.md)

