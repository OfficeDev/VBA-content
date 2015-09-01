
# Sentences Object (Word)

 **Last modified:** July 28, 2015

A collection of  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)**objects that represent all the sentences in a selection, range, or document. There is no Sentence object.

## Remarks

Use the  **Sentences** property to return the **Sentences** collection. The following example displays the number of sentences selected.


```
MsgBox Selection.Sentences.Count &amp; " sentences are selected"
```

Use  **Sentences**(Index), where Index is the index number, to return a  **Range** object that represents a sentence. The index number represents the position of a sentence in the **Sentences** collection. The following example formats the first sentence in the active document.




```
With ActiveDocument.Sentences(1) 
 .Bold = True 
 .Font.Size = 24 
End With
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

The  **Add** method isn't available for the **Sentences** collection. Instead, use the **InsertAfter**or  **InsertBefore**method to add a sentence to a  **Range** object. The following example inserts a sentence after the first paragraph in the active document.




```
With ActiveDocument 
 MsgBox .Sentences.Count &amp; " sentences" 
 .Paragraphs(1).Range.InsertParagraphAfter 
 .Paragraphs(2).Range.InsertBefore "The house is blue." 
 MsgBox .Sentences.Count &amp; " sentences" 
End With
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Sentences Object Members](a4668263-ff76-6f12-15f5-951d5db96431.md)
