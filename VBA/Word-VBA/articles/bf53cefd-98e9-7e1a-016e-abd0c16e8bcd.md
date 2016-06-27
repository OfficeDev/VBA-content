
# Document.EndReview Method (Word)

Terminates a review of a file that has been sent for review using the  **[SendForReview](2f2cdd5c-eeca-d03f-bd58-b5586f8f461f.md)** method or that has been automatically placed in a review cycle by sending a document to another user in an e-mail message.


## Syntax

 _expression_ . **EndReview**

 _expression_ Required. A variable that represents a **[Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


## Remarks

When executed, the  **EndReview** method displays a message asking the user whether to end the review.


## Example

This example terminates the review of the active document. This example assumes the active document part of a review cycle.


```vb
Sub EndDocRev() 
 ActiveDocument.EndReview 
End Sub
```


## See also


#### Concepts


[Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


[Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
