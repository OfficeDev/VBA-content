
# Sentences Object (Word)

A collection of  **[Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** objects that represent all the sentences in a selection, range, or document. There is no Sentence object.


## Remarks

Use the  **Sentences** property to return the **Sentences** collection. The following example displays the number of sentences selected.


```
MsgBox Selection.Sentences.Count &amp; " sentences are selected"
```

Use  **Sentences** (Index), where Index is the index number, to return a **Range** object that represents a sentence. The index number represents the position of a sentence in the **Sentences** collection. The following example formats the first sentence in the active document.




```
With ActiveDocument.Sentences(1) 
 .Bold = True 
 .Font.Size = 24 
End With
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

The  **Add** method isn't available for the **Sentences** collection. Instead, use the **InsertAfter** or **InsertBefore** method to add a sentence to a **Range** object. The following example inserts a sentence after the first paragraph in the active document.




```
With ActiveDocument 
 MsgBox .Sentences.Count &amp; " sentences" 
 .Paragraphs(1).Range.InsertParagraphAfter 
 .Paragraphs(2).Range.InsertBefore "The house is blue." 
 MsgBox .Sentences.Count &amp; " sentences" 
End With
```


## Methods



|**Name**|
|:-----|
|[Item](e68b4bac-c7b2-9953-d24d-e97e6b2f026c.md)|

## Properties



|**Name**|
|:-----|
|[Application](4549711b-1fa3-4296-a3cf-81506bea73f5.md)|
|[Count](e122ea1d-44e2-5f06-47e2-5058339efe0a.md)|
|[Creator](69465368-9258-cfc2-f469-69b27940e24e.md)|
|[First](4d9e4010-4aac-c060-285c-5a4665062874.md)|
|[Last](b116502a-ee26-934b-aa19-c589aafd90a0.md)|
|[Parent](e539a6c6-dade-b51f-e86e-cd68a24b9bd9.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)