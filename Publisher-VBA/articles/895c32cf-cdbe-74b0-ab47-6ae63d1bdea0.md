
# TextRange.Paragraphs Method (Publisher)

 **Last modified:** July 28, 2015

Returns a  ** [TextRange](566f240b-d2a6-8cb3-9eb7-68328d6c28bd.md)**object that represents the specified paragraphs.

## Syntax

 _expression_. **Paragraphs**( **_Start_**,  **_Length_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Start|Required| **Long**|The first paragraph in the returned range.|
|Length|Optional| **Long**|The number of paragraphs to be returned. Default is 1.|

### Return Value

TextRange


## Example

If  **_Length_** is omitted, the returned range contains one paragraph.



If  **_Length_** is greater than the number of paragraphs from the specified starting paragraph to the end of the text, the returned range contains all those paragraphs.

This example formats as indents the first line of the selected paragraph.




```
Sub FormatCurrentParagraph() 
 Selection.TextRange.Paragraphs(Start:=1).ParagraphFormat _ 
 .FirstLineIndent = InchesToPoints(0.5) 
End Sub
```

