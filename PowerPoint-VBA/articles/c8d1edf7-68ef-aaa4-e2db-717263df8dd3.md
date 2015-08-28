
# TextRange.Copy Method (PowerPoint)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Copies the specified object to the Clipboard.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Copy**

 _expression_A variable that represents a  **TextRange** object.


## Remarks
<a name="sectionSection1"> </a>

Use the  **Paste**method to paste the contents of the Clipboard.


## Example
<a name="sectionSection2"> </a>

This example copies the text in shape one on slide one in the active presentation to the Clipboard.


```
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Copy
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [TextRange Object](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)
#### Other resources


 [TextRange Object Members](cb8dc5ff-34de-3d04-1d56-ed387daaf6b9.md)
