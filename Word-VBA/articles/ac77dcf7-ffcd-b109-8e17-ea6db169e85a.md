
# Range.InsertBefore Method (Word)

Inserts the specified text before the specified range.


## Syntax

 _expression_ . **InsertBefore**( **_Text_** )

 _expression_ Required. A variable that represents a **[Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be inserted.|

## Remarks

After the text is inserted, the range is expanded to include the new text. If the range is a bookmark, the bookmark is also expanded to include the next text.

You can insert characters such as quotation marks, tab characters, and nonbreaking hyphens by using the Visual Basic  **Chr** function with the **InsertBefore** method. You can also use the following Visual Basic constants: **vbCr** , **vbLf** , **vbCrLf** and **vbTab** .


## Example

This example inserts the text "Introduction" as a separate paragraph at the beginning of the active document.


```vb
With ActiveDocument.Content 
 .InsertParagraphBefore 
 .InsertBefore "Introduction" 
End With
```


## See also


#### Concepts


[Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


[Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
