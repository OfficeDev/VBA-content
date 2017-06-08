---
title: Range.InsertBreak Method (Word)
keywords: vbawd10.chm157155450
f1_keywords:
- vbawd10.chm157155450
ms.prod: word
api_name:
- Word.Range.InsertBreak
ms.assetid: 9c565036-e060-f26e-2e12-9c340331233e
ms.date: 06/08/2017
---


# Range.InsertBreak Method (Word)

Inserts a page, column, or section break.


## Syntax

 _expression_ . **InsertBreak**( **_Type_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|The type of break to be inserted.Can be one of the  **WdBreakType** constants. If omitted, the default value is **wdPageBreak** .|

## Remarks

When you insert a page or column break, the range is replaced by the break. If you don't want to replace the range, use the  **Collapse** method before using the **InsertBreak** method. When you insert a section break, the break is inserted immediately preceding the **Range** .

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.


## Example

This example inserts a page break immediately following the second paragraph in the active document.


```vb
Set myRange = ActiveDocument.Paragraphs(2).Range 
With myRange 
 .Collapse Direction:=wdCollapseEnd 
 .InsertBreak Type:=wdPageBreak 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

