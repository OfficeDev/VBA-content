---
title: StyleSheet Object (Word)
keywords: vbawd10.chm2543
f1_keywords:
- vbawd10.chm2543
ms.prod: word
api_name:
- Word.StyleSheet
ms.assetid: 5e576ff8-c458-f5bd-730d-9db827c4f76e
ms.date: 06/08/2017
---


# StyleSheet Object (Word)

Represents a single cascading style sheet attached to a Web document. The  **StyleSheet** object is a member of the **[StyleSheets](stylesheets-object-word.md)** collection. The **StyleSheets** collection contains all the cascading style sheets attached to a specified document.


## Remarks

Use the  **Item** method or **StyleSheets** (Index), where Index is the name or number of the style sheet, of the **StyleSheets** collection to return a **StyleSheet** object. The following example removes the second style sheet from the **StyleSheets** collection.


```vb
Sub WebStyleSheets() 
 ActiveDocument.StyleSheets.Item(2).Delete 
End Sub
```

Use the  **Index** property to determine the precedence of cascading style sheets. The following example creates a table of attached cascading style sheets, ordered and indexed according to which style sheet is most important.




```vb
Sub CSSTable() 
 Dim styCSS As StyleSheet 
 
 With ActiveDocument.Range(Start:=0, End:=0) 
 .InsertAfter "CSS Name" &; vbTab &; "Index" 
 .InsertParagraphAfter 
 For Each styCSS In ActiveDocument.StyleSheets 
 .InsertAfter styCSS.Name &; vbTab &; styCSS.Index 
 .InsertParagraphAfter 
 Next styCSS 
 .ConvertToTable 
 End With 
End Sub
```

Use the  **Move** method to reorder the precedence of attached style sheets. The following example moves the most important style sheet to the least important of all attached cascading style sheets.




```vb
Sub MoveCSS() 
 ActiveDocument.StyleSheets(1) _ 
 .Move wdStyleSheetPrecedenceLowest 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


