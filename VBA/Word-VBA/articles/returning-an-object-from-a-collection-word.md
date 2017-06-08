---
title: Returning an Object from a Collection (Word)
ms.prod: word
ms.assetid: 28f76384-f495-9640-a7c8-10ada3fac727
ms.date: 06/08/2017
---


# Returning an Object from a Collection (Word)

The  **Item** method returns a single object from a collection. The following example sets the `docFirst` variable to a **[Document](document-object-word.md)** object that represents the first document in the **[Documents](documents-object-word.md)** collection.


```vb
Sub SetFirstDoc() 
    Dim docFirst As Document 
    Set docFirst = Documents.Item(1) 
End Sub
```


The  **Item** method is the default method for most collections, so you can write the same statement more concisely by omitting the **Item** keyword.




```vb
Sub SetFirstDoc() 
    Dim docFirst As Document 
    Set docFirst = Documents(1) 
End Sub
```


## Named Objects

Although you can usually specify an integer value with the  **Item** method, it may be more convenient to return an object by name. The following example switches the focus to a document named Sales.doc.


```vb
Sub ActivateDocument() 
    Documents("Sales.doc").Activate 
    MsgBox ActiveDocument.Name 
End Sub
```

The following example selects the text marked by the first bookmark in the active document.




```vb
Sub SelectBookmark() 
    ActiveDocument.Bookmarks(1).Select 
    MsgBox Selection.Text 
End Sub
```

Not all collections can be indexed by name. To determine the valid collection index values, see the collection object topic.


## Predefined Index Values

Some collections have predefined index values that you can use to return single objects. Each predefined index value is represented by a constant. For example, you specify a  **[WdBorderType](wdbordertype-enumeration-word.md)** constant with the **Borders** property to return a single **[Border](border-object-word.md)** object.

The following example adds a single 0.75-point border below the first paragraph in the selection.




```vb
Sub AddBorderToFirstParagraphInSelection() 
    With Selection.Paragraphs(1).Borders(wdBorderBottom) 
        .LineStyle = wdLineStyleSingle 
        .LineWidth = wdLineWidth300pt 
        .Color = wdColorBlue 
    End With 
End Sub
```


