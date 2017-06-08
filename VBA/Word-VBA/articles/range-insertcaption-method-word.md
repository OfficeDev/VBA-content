---
title: Range.InsertCaption Method (Word)
keywords: vbawd10.chm157155745
f1_keywords:
- vbawd10.chm157155745
ms.prod: word
api_name:
- Word.Range.InsertCaption
ms.assetid: fee41e81-1a78-2886-9693-dcf90da7c1bc
ms.date: 06/08/2017
---


# Range.InsertCaption Method (Word)

Inserts a caption immediately preceding or following the specified range.


## Syntax

 _expression_ . **InsertCaption**( **_Label_** , **_Title_** , **_TitleAutoText_** , **_Position_** , **_ExcludeLabel_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Label_|Required| **Variant**|The caption label to be inserted. Can be a  **String** or one of the **WdCaptionLabelID** constants. If the label has not yet been defined, an error occurs. Use the **Add** method with the **CaptionLabels** object to define new caption labels.|
| _Title_|Optional| **Variant**|The string to be inserted immediately following the label in the caption (ignored if TitleAutoText is specified).|
| _TitleAutoText_|Optional| **Variant**|The AutoText entry whose contents you want to insert immediately following the label in the caption (overrides any text specified by Title).|
| _Position_|Optional| **Variant**|Specifies whether the caption will be inserted above or below the range. Can be either one of the  **WdCaptionPosition** constants.|
| _ExcludeLabel_|Optional| **Variant**| **True** does not include the text label, as defined in the Label parameter. **False** includes the specified label.|

## Example

This example inserts a caption below the first table in the active document.


```vb
ActiveDocument.Tables(1).Range.InsertCaption _ 
 Label:=wdCaptionTable, _ 
 Position:=wdCaptionPositionBelow
```


## See also


#### Concepts


[Range Object](range-object-word.md)

