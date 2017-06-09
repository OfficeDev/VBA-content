---
title: Working with the Selection Object
ms.prod: word
ms.assetid: a1ef7e48-5a0f-d278-4b67-7b96f4e24052
ms.date: 06/08/2017
---


# Working with the Selection Object

When you work on a document in Word, you usually select text and then perform an action, such as formatting the text or typing text. In Visual Basic, it is usually not necessary to select text before modifying the text. Instead, you create a  **[Range](range-object-word.md)** object that refers to a specific portion of the document. For information about defining **Range** objects, see [Working with Range objects](working-with-range-objects.md). However, when you want your code to respond to or change a selection, you can do so by using the  **[Selection](selection-object-word.md)** object.

If text is not already selected, use the  **Select** method to select the text that is associated with a specific object and create a **Selection** object. For example, the following instruction selects the first word in the active document.



```vb
Sub SelectFirstWord() 
 ActiveDocument.Words(1).Select 
End Sub
```

For more information, see  [Selecting text in a document](selecting-text-in-a-document.md).
If text is already selected, use the  **[Selection](global-selection-property-word.md)** property to return a **Selection** object that represents the current selection in a document. There can be only one **Selection** object per document, and it always accesses the current selection. The following example changes the formatting of the paragraphs in the current selection.



```vb
Sub FormatSelection() 
 Selection.Paragraphs.LeftIndent = InchesToPoints(0.5) 
End Sub
```

This example inserts the word "Hello" after the current selection.



```vb
Sub InsertTextAfterSelection() 
 Selection.InsertAfter Text:="Hello " 
End Sub
```

This example applies bold formatting to the selected text.



```vb
Sub BoldSelectedText() 
 Selection.Font.Bold = True 
End Sub
```

The macro recorder often creates a macro that uses the  **Selection** object. The following example was created using the macro recorder. This macro selects the first two words in the active document and applies bold formatting to them.



```vb
Sub Macro() 
 Selection.HomeKey Unit:=wdStory 
 Selection.MoveRight Unit:=wdWord, Count:=2, Extend:=wdExtend 
 Selection.Font.Bold = wdToggle 
End Sub
```

The following example accomplishes the same task without selecting the text or using the  **Selection** object.



```vb
Sub WorkingWithRanges() 
 ActiveDocument.Range(Start:=0, _ 
 End:=ActiveDocument.Words(2).End).Bold = True 
End Sub
```


