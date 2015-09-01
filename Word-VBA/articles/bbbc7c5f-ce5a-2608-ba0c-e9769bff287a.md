
# Selection.Words Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Words](a718f69f-1db1-231a-9d65-bf20b48778ed.md)** collection that represents all the words in a selection. Read-only.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Words**

 _expression_A variable that represents a  ** [Selection](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)** object.


## Remarks
<a name="sectionSection1"> </a>

Punctuation and paragraph marks in a document are included in the  ** [Words](a718f69f-1db1-231a-9d65-bf20b48778ed.md)** collection. For information about returning a single member of a collection, see [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example displays the number of words in the selection. Paragraphs marks, partial words, and punctuation are included in the count.


```
MsgBox "There are " &amp; Selection.Words.Count &amp; " words."
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Selection Object](7b574a91-c33e-ecfd-6783-6b7528b2ed8f.md)
#### Other resources


 [Selection Object Members](71e67a43-d40a-ad9a-8ef2-c5c487733e0d.md)
