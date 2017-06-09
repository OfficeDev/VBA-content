---
title: Font.Duplicate Property (Word)
keywords: vbawd10.chm156368906
f1_keywords:
- vbawd10.chm156368906
ms.prod: word
api_name:
- Word.Font.Duplicate
ms.assetid: 86add1f8-9c1f-57c0-87d5-9fdef0841880
ms.date: 06/08/2017
---


# Font.Duplicate Property (Word)

Returns a copy of a **Font** object that represents the character formatting of the specified font.

## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

You can use the  **Duplicate** property to pick up the settings of all the properties of a duplicated **Font** object. You can assign the object returned by the **Duplicate** property to another **Font** object to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original.


## Example

This example sets the variable MyDupFont to the character formatting of the selection, removes bold formatting from MyDupFont, and adds italic formatting to it instead. The example also creates a new document, inserts text into it, and then applies the formatting stored in MyDupFont to the text.


```vb
Set myDupFont = Selection.Font.Duplicate 
With myDupFont 
 .Bold = False 
 .Italic = True 
End With 
Documents.Add 
Selection.InsertAfter "This is some text." 
Selection.Font = myDupFont
```


## See also


#### Concepts


[Font Object](font-object-word.md)

