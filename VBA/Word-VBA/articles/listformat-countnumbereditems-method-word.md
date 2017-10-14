---
title: ListFormat.CountNumberedItems Method (Word)
keywords: vbawd10.chm163578043
f1_keywords:
- vbawd10.chm163578043
ms.prod: word
api_name:
- Word.ListFormat.CountNumberedItems
ms.assetid: 71ce63e0-0e8d-53cc-178c-d91b0242482b
ms.date: 06/08/2017
---


# ListFormat.CountNumberedItems Method (Word)

Returns the number of bulleted or numbered items and LISTNUM fields in the specified  **ListFormat** object.


## Syntax

 _expression_ . **CountNumberedItems**

 _expression_ A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Example

This example formats the current selection as a list, using the second numbered list template. The example then counts the numbered and bulleted items and LISTNUM fields in the active document and displays the result in a message box.


```
Selection.Range.ListFormat.ApplyListTemplate _ 
    ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(2) 
Msgbox ActiveDocument.CountNumberedItems
```

This example counts the number of first-level numbered or bulleted items in the active document.




```
Msgbox ActiveDocument.Content.ListFormat _ 
    .CountNumberedItems(Level:=1)
```

This example counts the number of LISTNUM fields in the variable  _myRange_ . The result is displayed in a message box.






```vb
Set myDoc = ActiveDocumentSet myRange = _ 
    myDoc.Range(Start:=myDoc.Paragraphs(12).Range.Start, _ 
    End:=myDoc.Paragraphs(50).Range.End) 
numfields = myRange.ListFormat.CountNumberedItems(wdNumberListNum) 
Msgbox numfields
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

