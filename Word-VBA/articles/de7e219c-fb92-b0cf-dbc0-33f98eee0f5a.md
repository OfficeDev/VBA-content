
# ListFormat.ApplyNumberDefault Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Adds the default numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ApplyNumberDefault**( **_DefaultListBehavior_**)

 _expression_Required. A variable that represents a  ** [ListFormat](74773fd6-b713-34d4-b7be-f543c983008d.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultListBehavior|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants:  **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior**, but in new procedures you should use  **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.|

## Remarks
<a name="sectionSection1"> </a>

If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.


## Example
<a name="sectionSection2"> </a>

This example numbers the paragraphs in the selection. If the selection is already a numbered list, the example removes the numbers and formatting.


```
Selection.Range.ListFormat.ApplyNumberDefault
```

This example sets the variable myRange to include paragraphs three through six of the active document, and then it checks to see whether the range contains list formatting. If there is no list formatting, default numbers are applied to the range.




```
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyNumberDefault 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ListFormat Object](74773fd6-b713-34d4-b7be-f543c983008d.md)
#### Other resources


 [ListFormat Object Members](daf87b14-29a3-c5d9-ab43-8465237c02da.md)
