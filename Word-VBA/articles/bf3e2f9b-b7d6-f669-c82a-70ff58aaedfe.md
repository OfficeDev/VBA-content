
# Comments.Add Method (Word)

 **Last modified:** July 28, 2015

Returns a  **Comment** object that represents a comment added to a range.

## Syntax

 _expression_. **Add**( **_Range_**,  **_Text_**)

 _expression_Required. A variable that represents a  ** [Comments](e384b37a-50e3-a214-52a8-6fda2acc4991.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Range|Required| **Range object**|The range to have a comment added to it.|
|Text|Optional| **Variant**|The text of the comment.|

### Return Value

Comment


## Example

This example adds a comment at the insertion point.


```
Sub AddComment() 
 Selection.Collapse Direction:=wdCollapseEnd 
 ActiveDocument.Comments.Add _ 
 Range:=Selection.Range, Text:="review this" 
End Sub
```

This example adds a comment to the third paragraph in the active document.




```
Sub Comment3rd() 
 Dim myRange As Range 
 
 Set myRange = ActiveDocument.Paragraphs(3).Range 
 ActiveDocument.Comments.Add Range:=myRange, _ 
 Text:="original third paragraph" 
End Sub
```


## See also


#### Concepts


 [Comments Collection Object](e384b37a-50e3-a214-52a8-6fda2acc4991.md)
#### Other resources


 [Comments Object Members](2cd992bf-9e18-7f0e-3e8b-b3507ffd9bc7.md)
