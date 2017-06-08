---
title: Editing Text
ms.prod: word
ms.assetid: 55c4d9ee-00f1-fcc6-72b1-2c19f25420a6
ms.date: 06/08/2017
---


# Editing Text

This topic includes Visual Basic examples related to the following tasks:


-  [Determining whether text is selected](#Determining)
    
-  [Collapsing a selection or range](#Collapsing)
    
-  [Extending a selection or range](#Extending)
    
-  [Redefining a selection or range](#Redefining)
    
-  [Changing text](#Changing)
    

For information about, and examples of, other editing tasks, see the following topics:

 [Returning text from a document](returning-text-from-a-document.md)<br>
 [Selecting text in a document](selecting-text-in-a-document.md)<br>
 [Inserting text in a document](inserting-text-in-a-document.md)<br>
 [Manipulating a portion of a document](modifying-a-portion-of-a-document.md)<br>

## Determining whether text is selected

The  **[Type](selection-type-property-word.md)** property of the  **[Selection](selection-object-word.md)** object returns information about the type of selection. The following example displays a message if the selection is an insertion point.


```vb
Sub IsTextSelected() 
 If Selection.Type = wdSelectionIP Then MsgBox "Nothing is selected" 
End Sub
```


## Collapsing a selection or range

Use the  **Collapse**method to collapse a  **Selection** object or a **[Range](range-object-word.md)** object to its beginning or ending point. The following example collapses the selection to an insertion point at the beginning of the selection.


```vb
Sub CollapseToBeginning() 
 Selection.Collapse Direction:=wdCollapseStart 
End Sub
```

The following example cancels the range to its ending point (after the first word) and adds new text.




```vb
Sub CollapseToEnd() 
 Dim rngWords As Range 
 
 Set rngWords = ActiveDocument.Words(1) 
 With rngWords 
 .Collapse Direction:=wdCollapseEnd 
 .Text = "(This is a test.) " 
 End With 
End Sub
```


## Extending a selection or range

The following example uses the  **[MoveEnd](selection-moveend-method-word.md)** method of the  **Selection** object to extend the end of the selection to include three additional words. The **[MoveLeft](selection-moveleft-method-word.md)**,  **[MoveRight](selection-moveright-method-word.md)**,  **[MoveUp](selection-moveup-method-word.md)**, and  **[MoveDown](selection-movedown-method-word.md)** methods can also be used to extend a  **Selection** object.


```vb
Sub ExtendSelection() 
 Selection.MoveEnd Unit:=wdWord, Count:=3 
End Sub
```

The following example uses the  **[MoveEnd](range-moveend-method-word.md)** method of the **[Range](range-object-word.md)** object to extend the range to include the first three paragraphs in the active document.




```vb
Sub ExtendRange() 
 Dim rngParagraphs As Range 
 
 Set rngParagraphs = ActiveDocument.Paragraphs(1).Range 
 rngParagraphs.MoveEnd Unit:=wdParagraph, Count:=2 
End Sub
```


## Redefining a selection or range

Use the  **SetRange**method to redefine an existing  **Selection** object or **Range** object. For more information, see [Working with the Selection object](working-with-the-selection-object.md) or [Working with Range objects](working-with-range-objects.md).


## Changing text

You can change existing text by changing the contents of a range. The following instruction changes the first word in the active document by setting the  **[Text](range-text-property-word.md)** property of a  **Range** object to "The ".


```vb
Sub ChangeText() 
 ActiveDocument.Words(1).Text = "The " 
End Sub
```

You can also use the  **Delete**method or the  **Selection** object or the **Range** object to delete existing text, and then insert new text using the **InsertAfter**method or the  **InsertBefore**method. The following example deletes the first paragraph in the active document and inserts new text.




```vb
Sub DeleteText() 
 Dim rngFirstParagraph As Range 
 
 Set rngFirstParagraph = ActiveDocument.Paragraphs(1).Range 
 With rngFirstParagraph 
 .Delete 
 .InsertAfter Text:="New text" 
 .InsertParagraphAfter 
 End With 
End Sub
```


