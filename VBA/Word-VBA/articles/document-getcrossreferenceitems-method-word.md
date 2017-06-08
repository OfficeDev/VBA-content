---
title: Document.GetCrossReferenceItems Method (Word)
keywords: vbawd10.chm158007443
f1_keywords:
- vbawd10.chm158007443
ms.prod: word
api_name:
- Word.Document.GetCrossReferenceItems
ms.assetid: 380e3019-2574-f50b-d871-dcebb564b06e
ms.date: 06/08/2017
---


# Document.GetCrossReferenceItems Method (Word)

Returns an array of items that can be cross-referenced based on the specified cross-reference type.


## Syntax

 _expression_ . **GetCrossReferenceItems**( **_ReferenceType_** )

 _expression_ An expression that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReferenceType_|Required| **Variant**|The type of item you want to insert a cross-reference to. Can be any  **[WdReferenceType](wdreferencetype-enumeration-word.md)** constant.|

## Remarks

The array that this method returns corresponds to the items listed in the  **For which** box in the **Cross-reference** dialog box. The value returned by this method can be used as the value of the ReferenceWhich argument for the **InsertCrossReference** method of the **Range** or **Selection** object.


## Example

This example displays the name of the first bookmark in the active document that can be cross-referenced.


```vb
If ActiveDocument.Bookmarks.Count >= 1 Then 
 myBookmarks = ActiveDocument.GetCrossReferenceItems( _ 
 wdRefTypeBookmark) 
 MsgBox myBookmarks(1) 
End If
```

This example uses the  **GetCrossReferenceItems** method to retrieve a list of headings that can be cross-referenced and then inserts a cross-reference to the page that includes the heading "Introduction."




```vb
myHeadings = _ 
 ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading) 
For i = 1 To Ubound(myHeadings) 
 If Instr(LCase$(myHeadings(i)), "introduction") Then 
 Selection.InsertCrossReference _ 
 ReferenceType:=wdRefTypeHeading, _ 
 ReferenceKind:=wdPageNumber, ReferenceItem:=i 
 Selection.InsertParagraphAfter 
 End If 
Next i
```


## See also


#### Concepts


[Document Object](document-object-word.md)

