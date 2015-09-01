
# Document.ManualHyphenation Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Initiates manual hyphenation of a document, one line at a time.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ManualHyphenation**

 _expression_Required. A variable that represents a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


## Remarks
<a name="sectionSection1"> </a>

When you use the  **ManualHyphenation** method, Word prompts he user to accept or decline suggested hyphenations.


## Example
<a name="sectionSection2"> </a>

This example starts manual hyphenation of the active document.


```
ActiveDocument.ManualHyphenation
```

This example sets hyphenation options and then starts manual hyphenation of MyDoc.doc.




```
With Documents("MyDoc.doc") 
 .HyphenationZone = InchesToPoints(0.25) 
 .HyphenateCaps = False 
 .ManualHyphenation 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


 [Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
