---
title: Row.Select Method (Word)
keywords: vbawd10.chm156303359
f1_keywords:
- vbawd10.chm156303359
ms.prod: word
api_name:
- Word.Row.Select
ms.assetid: f3c31e32-b316-abf2-fec6-b76e8950b1b5
ms.date: 06/08/2017
---


# Row.Select Method (Word)

Selects the specified table row.


## Syntax

 _expression_ . **Select**

 _expression_ Required. A variable that represents a **[Row](row-object-word.md)** object.


## Remarks

After using this method, use the  **Selection** object to work with the selected row. For more information, see[Working with the Selection Object](http://msdn.microsoft.com/library/a1ef7e48-5a0f-d278-4b67-7b96f4e24052%28Office.15%29.aspx).


## Example

This example selects row one in table one of Report.doc.


```
Documents("Report.doc").Tables(1).Rows(1).Select
```


## See also


#### Concepts


[Row Object](row-object-word.md)

