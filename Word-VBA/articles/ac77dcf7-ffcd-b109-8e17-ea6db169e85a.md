
# Range.InsertBefore Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Inserts the specified text before the specified range.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **InsertBefore**( **_Text_**)

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Text|Required| **String**|The text to be inserted.|

## Remarks
<a name="sectionSection1"> </a>

After the text is inserted, the range is expanded to include the new text. If the range is a bookmark, the bookmark is also expanded to include the next text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertBefore** method. You can also use the following Visual Basic constants: **vbCr**,  **vbLf**,  **vbCrLf** and **vbTab**.


## Example
<a name="sectionSection2"> </a>

This example inserts the text "Introduction" as a separate paragraph at the beginning of the active document.


```
With ActiveDocument.Content 
 .InsertParagraphBefore 
 .InsertBefore "Introduction" 
End With
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
