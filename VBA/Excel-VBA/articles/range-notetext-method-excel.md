---
title: Range.NoteText Method (Excel)
keywords: vbaxl10.chm144166
f1_keywords:
- vbaxl10.chm144166
ms.prod: excel
api_name:
- Excel.Range.NoteText
ms.assetid: cd0e5073-7d04-a52c-f375-f7c59bc8f88a
ms.date: 06/08/2017
---


# Range.NoteText Method (Excel)

Returns or sets the cell note associated with the cell in the upper-left corner of the range. Read/write  **String** . Cell notes have been replaced by range comments. For more information, see the **[Comment](comment-object-excel.md)** object.


## Syntax

 _expression_ . **NoteText**( **_Text_** , **_Start_** , **_Length_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The text to add to the note (up to 255 characters). The text is inserted starting at position  _Start_, replacing  _Length_ characters of the existing note. If this argument is omitted, this method returns the current text of the note starting at position _Start_, for  _Length_ characters.|
| _Start_|Optional| **Variant**|The starting position for the text that?s set or returned. If this argument is omitted, this method starts at the first character. To append text to the note, specify a number larger than the number of characters in the existing note.|
| _Length_|Optional| **Variant**|The number of characters to be set or returned. If this argument is omitted, Microsoft Excel sets or returns characters from the starting position to the end of the note (up to 255 characters). If there are more than 255 characters from  _Start_ to the end of the note, this method returns only 255 characters.|

### Return Value

String


## Remarks

To add a note that contains more than 255 characters, use this method once to specify the first 255 characters, and then use it again to append the remainder of the note (no more than 255 characters at a time).


## Example

This example sets the cell note text for cell A1 on Sheet1.


```vb
Worksheets("Sheet1").Range("A1").NoteText "This may change!"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

