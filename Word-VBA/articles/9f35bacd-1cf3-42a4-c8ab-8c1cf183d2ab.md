
# Range.InsertFile Method (Word)

 **Last modified:** July 28, 2015

Inserts all or part of the specified file.

## Syntax

 _expression_. **InsertFile**( **_FileName_**,  **_Range_**,  **_ConfirmConversions_**,  **_Link_**,  **_Attachment_**)

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FileName|Required| **String**|The path and file name of the file to be inserted. If you don't specify a path, Word assumes the file is in the current folder.|
|Range|Optional| **Variant**|If the specified file is a Word document, this parameter refers to a bookmark. If the file is another type (for example, a Microsoft Excel worksheet), this parameter refers to a named range or a cell range (for example, R1C1:R3C4).|
|ConfirmConversions|Optional| **Variant**| **True** to have Word prompt you to confirm conversion when inserting files in formats other than the Word Document format.|
|Link|Optional| **Variant**| **True** to insert the file by using an INCLUDETEXT field.|
|Attachment|Optional| **Variant**| **True** to insert the file as an attachment to an e-mail message.|

## Example

This example uses an INCLUDETEXT field to insert the TEST.DOC file at the end of the current document.


```
ActiveDocument.Range.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Range.InsertFile FileName:="C:\TEST.DOC", Link:=True
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
