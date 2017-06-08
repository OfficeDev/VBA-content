---
title: ParagraphFormat.Duplicate Property (Word)
keywords: vbawd10.chm156434442
f1_keywords:
- vbawd10.chm156434442
ms.prod: word
api_name:
- Word.ParagraphFormat.Duplicate
ms.assetid: cc5e9633-ea7c-8317-5321-c7bbf1288579
ms.date: 06/08/2017
---


# ParagraphFormat.Duplicate Property (Word)

Returns a read-only  **ParagraphFormat** object that represents the paragraph formatting of the specified paragraph.


## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[ParagraphFormat](paragraphformat-object-word.md)** object.


## Remarks

You can use the  **Duplicate** property to pick up the settings of all the properties of a duplicated **ParagraphFormat** object. You can assign the object returned by the **Duplicate** property to another object of the same type to apply those settings all at once. Before assigning the duplicate object to another object, you can change any of the properties of the duplicate object without affecting the original.


## Example

This example duplicates the paragraph formatting of the first paragraph in the active document and stores the formatting in the variable myDup, and then it changes the left indent for myDup to 1 inch. The example also creates a new document, inserts text into it, and then applies the paragraph formatting stored in myDup to the text.


```vb
ActiveDocument.Range(Start:=0, End:=0).InsertAfter _ 
 "Paragraph Number 1" 
Set myDup = ActiveDocument.Paragraphs(1).Format.Duplicate 
myDup.LeftIndent = InchesToPoints(1) 
Documents.Add 
Selection.InsertAfter "This is a new paragraph." 
Selection.Paragraphs.Format = myDup
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-word.md)

