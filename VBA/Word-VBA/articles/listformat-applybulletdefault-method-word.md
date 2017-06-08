---
title: ListFormat.ApplyBulletDefault Method (Word)
keywords: vbawd10.chm163578068
f1_keywords:
- vbawd10.chm163578068
ms.prod: word
api_name:
- Word.ListFormat.ApplyBulletDefault
ms.assetid: 40e0b8f6-9360-441b-a7fc-52bff8953ea8
ms.date: 06/08/2017
---


# ListFormat.ApplyBulletDefault Method (Word)

Adds bullets and formatting to the paragraphs in the range for the specified  **ListFormat** object.


## Syntax

 _expression_ . **ApplyBulletDefault**( **_DefaultListBehavior_** )

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following constants:  **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior** , but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting with respect to indenting and multilevel lists.|

## Remarks

If the paragraphs are already formatted with bullets, this method removes the bullets and formatting.


## Example

This example adds bullets and formatting to the paragraphs in the selection. If there are already bullets in the selection, the example removes the bullets and formatting.


```
Selection.Range.ListFormat.ApplyBulletDefault
```

This example adds a bullet and formatting to, or removes them from, the second paragraph in MyDoc.doc.




```
Documents("MyDoc.doc").Paragraphs(2).Range.ListFormat _ 
 .ApplyBulletDefault
```

This example sets the variable myRange to a range that includes paragraphs three through six of the active document, and then it checks to see whether the range contains list formatting. If there is no list formatting, default bullets are added.




```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyBulletDefault 
End If
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

