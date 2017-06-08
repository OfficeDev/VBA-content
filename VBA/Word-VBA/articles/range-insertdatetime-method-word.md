---
title: Range.InsertDateTime Method (Word)
keywords: vbawd10.chm157155772
f1_keywords:
- vbawd10.chm157155772
ms.prod: word
api_name:
- Word.Range.InsertDateTime
ms.assetid: 2203a0bb-6c90-ee55-6bdc-73f6761e4603
ms.date: 06/08/2017
---


# Range.InsertDateTime Method (Word)

Inserts the current date or time, or both, either as text or as a TIME field.


## Syntax

 _expression_ . **InsertDateTime**( **_DateTimeFormat_** , **_InsertAsField_** , **_InsertAsFullWidth_** , **_DateLanguage_** , **_CalendarType_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DateTimeFormat_|Optional| **Variant**|The format to be used for displaying the date or time, or both. If this argument is omitted, Microsoft Word uses the short-date style from the Windows Control Panel ( **Regional Settings** icon).|
| _InsertAsField_|Optional| **Variant**| **True** to insert the specified information as a TIME field. The default value is **True** .|
| _InsertAsFullWidth_|Optional| **Variant**| **True** to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _DateLanguage_|Optional| **Variant**|Sets the language in which to display the date or time. Can be either of the  **WdDateLanguage** constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _CalendarType_|Optional| **Variant**|Sets the calendar type to use when displaying the date or time. Can be either of the  **WdCalendarTypeBi** constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|

## Example

This example inserts the current date at the end of the active document. A possible result might be "01/12/99."


```vb
With ActiveDocument.Content 
 .Collapse Direction:=wdCollapseEnd 
 .InsertDateTime DateTimeFormat:="MM/dd/yy", _ 
 InsertAsField:=False 
End With
```

This example inserts a TIME field for the current date in the footer for the active document.




```vb
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range _ 
 .InsertDateTime DateTimeFormat:="MMMM dd, yyyy", _ 
 InsertAsField:=True
```


## See also


#### Concepts


[Range Object](range-object-word.md)

