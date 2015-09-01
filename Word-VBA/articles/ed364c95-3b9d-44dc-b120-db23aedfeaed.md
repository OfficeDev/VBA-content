
# Selection.Shrink Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Shrinks the selection to the next smaller unit of text.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Shrink**

 _expression_A variable that represents a  ** [Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks
<a name="sectionSection1"> </a>

The unit progression for this method is as follows: entire document, section, paragraph, sentence, word, insertion point.


## Example
<a name="sectionSection2"> </a>

This example collapses the selected text to the next smaller unit of text.


```
If Selection.Type = wdSelectionNormal Then 
 Selection.Shrink 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


 [Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
