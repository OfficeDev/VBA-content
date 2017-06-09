---
title: Table.Title Property (Word)
keywords: vbawd10.chm156303569
f1_keywords:
- vbawd10.chm156303569
ms.prod: word
api_name:
- Word.Table.Title
ms.assetid: a7b8437a-3882-1301-4235-7491156aca3a
ms.date: 06/08/2017
---


# Table.Title Property (Word)

Returns or sets a  **String** that contains a title for the specified table. Read/write.


## Syntax

 _expression_ . **Title**

 _expression_ A variable that represents a **[Table](table-object-word.md)** object.


## Remarks

Use the  **Title** property to provide an alternative text title for a table. This property adds title text to the **Title** text box on the **Alt Text** tab of the **Table Properties** dialog in Word.


 **Note**  Web browsers display alternative text while tables are loading or if they are missing. Web search engines use the alternative text to help find Web pages. Alternative text is also used to assist disabilities.


## Example

The following code example adds an alternative text table title to the first table in the active document.


```vb
ActiveDocument.Tables(1).Title = "Table 1."
```


## See also


#### Concepts


[Table Object](table-object-word.md)

