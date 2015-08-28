
# ParagraphFormat.HangingPunctuation Property (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets the hanging punctuation option if you have an Asian language setting specified. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **HangingPunctuation**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks
<a name="sectionSection1"> </a>

The value of the  **HangingPunctuation** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The hanging punctuation option is not selected.|
| **msoTrue**| The hanging punctuation option is selected.|

## Example
<a name="sectionSection2"> </a>

This example selects hanging punctuation for the first paragraph of the active presentation.


```
ActivePresentation.Paragraphs(1).HangingPunctuation = msoTrue
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ParagraphFormat Object](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)
#### Other resources


 [ParagraphFormat Object Members](c269be7c-ad65-672d-bcac-2a4913346d3e.md)
