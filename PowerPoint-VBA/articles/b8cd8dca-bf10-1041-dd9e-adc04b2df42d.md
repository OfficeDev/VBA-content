
# TextRange.Words Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents the specified subset of text words.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Words**( **_Start_**,  **_Length_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Start|Optional| **Long**|The first word in the returned range.|
|Length|Optional| **Long**|The number of words to be returned.|

### Return Value

TextRange


## Remarks
<a name="sectionSection1"> </a>

For information about counting or looping through the words in a text range, see the  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object.

If both Start and Length are omitted, the returned range starts with the first word and ends with the last paragraph in the specified range.

If Start is specified but Length is omitted, the returned range contains one word.

If Length is specified but Start is omitted, the returned range starts with the first word in the specified range.

If Start is greater than the number of words in the specified text, the returned range starts with the last word in the specified range.

If Length is greater than the number of words from the specified starting word to the end of the text, the returned range contains all those words.


## Example
<a name="sectionSection2"> </a>

This example formats as bold the second, third, and fourth words in the first paragraph in shape two on slide one in the active presentation.


```
Application.ActivePresentation.Slides(1).Shapes(2) _

    .TextFrame.TextRange.Paragraphs(1).Words(2, 3).Font _

    .Bold = True


```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [TextRange Object](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)
#### Other resources


 [TextRange Object Members](cb8dc5ff-34de-3d04-1d56-ed387daaf6b9.md)
