---
title: Application.DefaultTableSeparator Property (Word)
keywords: vbawd10.chm158335081
f1_keywords:
- vbawd10.chm158335081
ms.prod: word
api_name:
- Word.Application.DefaultTableSeparator
ms.assetid: eb393e87-c408-8911-a1e3-8f04e5ce66c6
ms.date: 06/08/2017
---


# Application.DefaultTableSeparator Property (Word)

Returns or sets the single character used to separate text into cells when text is converted to a table. Read/write  **String** .


## Syntax

 _expression_ . **DefaultTableSeparator**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The value of the  **DefaultTableSeparator** property is used if the Separator argument is omitted from the **ConvertToTable** method or the **[Range](range-object-word.md)** or **[Selection](selection-object-word.md)** object.


## Example

This example changes the default table separator character.


```vb
Application.DefaultTableSeparator = "%"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

