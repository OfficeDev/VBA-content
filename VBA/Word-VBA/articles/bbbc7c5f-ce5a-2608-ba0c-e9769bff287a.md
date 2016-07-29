
# Selection.Words Property (Word)

Returns a  **[Words](a718f69f-1db1-231a-9d65-bf20b48778ed.md)** collection that represents all the words in a selection. Read-only.


## Syntax

 _expression_ . **Words**

 _expression_ A variable that represents a **[Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks

Punctuation and paragraph marks in a document are included in the  **[Words](a718f69f-1db1-231a-9d65-bf20b48778ed.md)** collection. For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of words in the selection. Paragraphs marks, partial words, and punctuation are included in the count.


```vb
MsgBox "There are " &; Selection.Words.Count &; " words."
```


## See also


#### Concepts


[Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


[Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
