
# ListFormat.ApplyNumberDefault Method (Word)

Adds the default numbering scheme to the paragraphs in the range for the specified  **ListFormat** object.


## Syntax

 _expression_ . **ApplyNumberDefault**( **_DefaultListBehavior_** )

 _expression_ Required. A variable that represents a **[ListFormat](74773fd6-b713-34d4-b7be-f543c983008d.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants:  **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior** , but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.|

## Remarks

If the paragraphs are already formatted as a numbered list, this method removes the numbers and formatting.


## Example

This example numbers the paragraphs in the selection. If the selection is already a numbered list, the example removes the numbers and formatting.


```
Selection.Range.ListFormat.ApplyNumberDefault
```

This example sets the variable myRange to include paragraphs three through six of the active document, and then it checks to see whether the range contains list formatting. If there is no list formatting, default numbers are applied to the range.




```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyNumberDefault 
End If
```


## See also


#### Concepts


[ListFormat Object](74773fd6-b713-34d4-b7be-f543c983008d.md)
#### Other resources


[ListFormat Object Members](daf87b14-29a3-c5d9-ab43-8465237c02da.md)
