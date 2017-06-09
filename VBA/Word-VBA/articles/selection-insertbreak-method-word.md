---
title: Selection.InsertBreak Method (Word)
keywords: vbawd10.chm158662778
f1_keywords:
- vbawd10.chm158662778
ms.prod: word
api_name:
- Word.Selection.InsertBreak
ms.assetid: 2c9d8cb8-1cc1-3d69-1e26-3a6878c0b1da
ms.date: 06/08/2017
---


# Selection.InsertBreak Method (Word)

Inserts a page, column, or section break.


## Syntax

 _expression_ . **InsertBreak**( **_Type_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[WdBreakType](wdbreaktype-enumeration-word.md)**|the type of break to insert. The default value is  **wdPageBreak** . Some of the **WdBreakType** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.|

## Remarks

When you insert a page or column break, the break replaces the selection. If you don't want to replace the selection, use the  **[Collapse](selection-collapse-method-word.md)** method before using the **InsertBreak** method.


 **Note**  When you insert a section break, the break is inserted immediately preceding the selection.


## Example

This example inserts a continuous section break immediately preceding the selection.


```
Selection.InsertBreak Type:=wdSectionBreakContinuous
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

