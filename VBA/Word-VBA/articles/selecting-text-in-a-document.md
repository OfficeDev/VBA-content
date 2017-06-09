---
title: Selecting Text in a Document
ms.prod: word
ms.assetid: 1684b81f-caed-ea76-1378-580f6e34a1db
ms.date: 06/08/2017
---


# Selecting Text in a Document

Use the  **Select**method to select an item in a document. The  **Select** method is available from several objects, such as **[Bookmark](bookmark-object-word.md)**,  **[Field](field-object-word.md)**,  **[Range](range-object-word.md)**, and  **[Table](table-object-word.md)**. The following example selects the first table in the active document.


```vb
Sub SelectTable() 
 ActiveDocument.Tables(1).Select 
End Sub
```


The following example selects the first field in the active document.




```vb
Sub SelectField() 
 ActiveDocument.Fields(1).Select 
End Sub
```

The following example selects the first four paragraphs in the active document. The  **Range**method is used to create a  **Range** object that refers to the first four paragraphs. The **Select** method is then applied to the **Range** object.



```vb
Sub SelectRange() 
 Dim rngParagraphs As Range 
 Set rngParagraphs = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(1).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End) 
 rngParagraphs.Select 
End Sub
```

For more information, see  [Working with the Selection object](working-with-the-selection-object.md).

