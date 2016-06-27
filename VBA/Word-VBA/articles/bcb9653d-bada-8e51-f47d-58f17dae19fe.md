
# Sentences Object (Word)

A collection of  **[Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** objects that represent all the sentences in a selection, range, or document. There is no Sentence object.


## Remarks

Use the  **Sentences** property to return the **Sentences** collection. The following example displays the number of sentences selected.


```vb
MsgBox Selection.Sentences.Count &; " sentences are selected"
```

Use  **Sentences** (Index), where Index is the index number, to return a **Range** object that represents a sentence. The index number represents the position of a sentence in the **Sentences** collection. The following example formats the first sentence in the active document.




```vb
With ActiveDocument.Sentences(1) 
 .Bold = True 
 .Font.Size = 24 
End With
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

The  **Add** method isn't available for the **Sentences** collection. Instead, use the **InsertAfter** or **InsertBefore** method to add a sentence to a **Range** object. The following example inserts a sentence after the first paragraph in the active document.




```vb
With ActiveDocument 
 MsgBox .Sentences.Count &; " sentences" 
 .Paragraphs(1).Range.InsertParagraphAfter 
 .Paragraphs(2).Range.InsertBefore "The house is blue." 
 MsgBox .Sentences.Count &; " sentences" 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
[Sentences Object Members](a4668263-ff76-6f12-15f5-951d5db96431.md)
