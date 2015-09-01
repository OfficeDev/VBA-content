
# LetterContent.AttentionLine Property (Word)

 **Last modified:** July 28, 2015

Returns or sets the attention line text for a letter created by the Letter Wizard. Read/write  **String**.

## Syntax

 _expression_. **AttentionLine**

 _expression_A variable that represents a  ** [LetterContent](62a4e17a-6598-c904-f27d-817c19c04981.md)** object.


## Example

This example retrieves the Letter Wizard elements from the active document. If the attention line isn't blank, the example displays the text in a message box.


```
If ActiveDocument.GetLetterContent.AttentionLine <> "" Then 
 MsgBox ActiveDocument.GetLetterContent.AttentionLine 
End If
```

This example retrieves the Letter Wizard elements from the active document, changes the attention line text, and then uses the  ** [SetLetterContent](8c9b2f6e-34a7-41a3-761d-c1a5da141aba.md)** method to update the document to reflect the changes.




```
Dim lcTemp As LetterContent 
 
Set lcTemp = ActiveDocument.GetLetterContent 
 
lcTemp.AttentionLine = "Greetings" 
ActiveDocument.SetLetterContent LetterContent:=lcTemp
```


## See also


#### Concepts


 [LetterContent Object](62a4e17a-6598-c904-f27d-817c19c04981.md)
#### Other resources


 [LetterContent Object Members](614f0a71-9722-0847-5b5f-fd6b0a85bd2f.md)
