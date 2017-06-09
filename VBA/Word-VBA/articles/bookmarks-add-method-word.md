---
title: Bookmarks.Add Method (Word)
keywords: vbawd10.chm157745157
f1_keywords:
- vbawd10.chm157745157
ms.prod: word
api_name:
- Word.Bookmarks.Add
ms.assetid: 647795da-d7e2-7b6f-c412-5b684ec962a2
ms.date: 06/08/2017
---


# Bookmarks.Add Method (Word)

Returns a  **Bookmark** object that represents a bookmark added to a range.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Range_** )

 _expression_ Required. A variable that represents a **[Bookmarks](bookmarks-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the bookmark. The name cannot be more than 40 characters or include more than one word.|
| _Range_|Optional| **Variant**|The range of text marked by the bookmark. A bookmark can be set to a collapsed range (the insertion point).|

### Return Value

Bookmark


## Example

This example adds a bookmark named myplace to the selection in the active document.


```vb
Sub BMark() 
 ' Select some text in the active document prior 
 ' to execution. 
 ActiveDocument.Bookmarks.Add _ 
 Name:="myplace", Range:=Selection.Range 
End Sub
```

This example adds a bookmark named mark at the insertion point.




```vb
Sub Mark() 
 ActiveDocument.Bookmarks.Add Name:="mark" 
End Sub
```

This example adds a bookmark named third_para to the third paragraph in Letter.doc, and then it displays all the bookmarks for the document in the active window.




```vb
Sub ThirdPara() 
 Dim myDoc As Document 
 
 ' To best illustrate this example, 
 ' Letter.doc must be opened, not active, 
 ' and contain more than 3 paragraphs. 
 Set myDoc = Documents("Letter.doc") 
 myDoc.Bookmarks.Add Name:="third_para", _ 
 Range:=myDoc.Paragraphs(3).Range 
 myDoc.ActiveWindow.View.ShowBookmarks = True 
End Sub
```


## See also


#### Concepts


[Bookmarks Collection Object](bookmarks-object-word.md)

