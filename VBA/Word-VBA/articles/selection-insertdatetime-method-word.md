---
title: Selection.InsertDateTime Method (Word)
keywords: vbawd10.chm158663100
f1_keywords:
- vbawd10.chm158663100
ms.prod: word
api_name:
- Word.Selection.InsertDateTime
ms.assetid: f9cfca41-e0f2-4656-5fa2-2463c50af1f5
ms.date: 06/08/2017
---


# Selection.InsertDateTime Method (Word)

Inserts the current date or time, or both, either as text or as a TIME field.


## Syntax

 _expression_ . **InsertDateTime**( **_DateTimeFormat_** , **_InsertAsField_** , **_InsertAsFullWidth_** , **_DateLanguage_** , **_CalendarType_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DateTimeFormat_|Optional| **Variant**|The format to be used for displaying the date or time, or both. If this argument is omitted, Microsoft Word uses the short-date style from the Windows Control Panel ( **Regional Settings** icon).|
| _InsertAsField_|Optional| **Variant**| **True** to insert the specified information as a TIME field. The default value is **True** .|
| _InsertAsFullWidth_|Optional| **Variant**| **True** to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _DateLanguage_|Optional| **Variant**|Sets the language in which to display the date or time. Can be either of the  **WdDateLanguage** constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|
| _CalendarType_|Optional| **Variant**|Sets the calendar type to use when displaying the date or time. Can be either of the  **WdCalendarTypeBi** constants. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|

## Remarks

Using this method replaces the current selection. To avoid this, use the  **[Collapse](selection-collapse-method-word.md)** method before using this method.


## Example

This example inserts a TIME field for the current date. A possible result might be "November 18, 1999."


```
Selection.InsertDateTime DateTimeFormat:="MMMM dd, yyyy", _ 
 InsertAsField:=True
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

