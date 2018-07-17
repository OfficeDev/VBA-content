---
title: Range.InsertCrossReference Method (Word)
keywords: vbawd10.chm157155746
f1_keywords:
- vbawd10.chm157155746
ms.prod: word
api_name:
- Word.Range.InsertCrossReference
ms.assetid: 5899db5b-254c-17ac-4c4b-943a5a5b44cb
ms.date: 06/08/2017
---


# Range.InsertCrossReference Method (Word)

Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table).


## Syntax

 _expression_ . **InsertCrossReference**( **_ReferenceType_** , **_ReferenceKind_** , **_ReferenceItem_** , **_InsertAsHyperlink_** , **_IncludePosition_** , **_SeparateNumbers_** , **_SeparatorString_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReferenceType_|Required| **Variant**|The type of item for which a cross-reference is to be inserted. Can be any  **WdReferenceType** or **WdCaptionLabelID** constant or a user defined caption label.|
| _ReferenceKind_|Required| **WdReferenceKind**|The information to be included in the cross-reference.|
| _ReferenceItem_|Required| **Variant**|If ReferenceType is  **wdRefTypeBookmark** , this argument specifies a bookmark name. For all other ReferenceType values, this argument specifies the item number or name in the **Reference type** option in the **Cross-reference** dialog box. Use the **GetCrossReferenceItems** method to return a list of item names that can be used with this argument.|
| _InsertAsHyperlink_|Optional| **Variant**| **True** to insert the cross-reference as a hyperlink to the referenced item.|
| _IncludePosition_|Optional| **Variant**| **True** to insert "above" or "below," depending on the location of the reference item in relation to the cross-reference.|
| _SeparateNumbers_|Optional| **Variant**| **True** to use a separator to separate the numbers from the associated text. (Use only if the ReferenceType parameter is set to **wdRefTypeNumberedItem** and the ReferenceKind parameter is set to **wdNumberFullContext** .)|
| _SeparatorString_|Optional| **Variant**|Specifies the string to use as a separator if the SeparateNumbers parameter is set to  **True** .|

## Remarks

If you specify  **wdPageNumber** for the value of ReferenceKind, you may need to repaginate the document to see the correct cross-reference information.


## Example

This example inserts at the beginning of the active document a cross-reference to the page that includes the first bookmark in the document.


```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
myBookmarks = ActiveDocument _ 
 .GetCrossReferenceItems(wdRefTypeBookmark) 
With myRange 
 .InsertBefore "Page " 
 .Collapse Direction:=wdCollapseEnd 
 .InsertCrossReference ReferenceType:=wdRefTypeBookmark, _ 
 ReferenceKind:=wdPageNumber, ReferenceItem:=myBookmarks(1) 
End With
```


## See also


#### Concepts


[Range Object](range-object-word.md)

