---
title: AutoCaption.AutoInsert Property (Word)
keywords: vbawd10.chm159055873
f1_keywords:
- vbawd10.chm159055873
ms.prod: word
api_name:
- Word.AutoCaption.AutoInsert
ms.assetid: eac9cee8-93d5-e707-b03d-ef1dbe906ef9
ms.date: 06/08/2017
---


# AutoCaption.AutoInsert Property (Word)

 **True** if a caption is automatically added when the item is inserted into a document. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoInsert**

 _expression_ A variable that represents an **[AutoCaption](autocaption-object-word.md)** object.


## Example

This example enables Word to add captions to tables automatically. Then the example collapses the selection to an insertion point, and inserts a table. A caption is automatically added to the new table.


```vb
AutoCaptions("Microsoft Word Table").AutoInsert = True 
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.Tables.Add Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=2
```


## See also


#### Concepts


[AutoCaption Object](autocaption-object-word.md)

