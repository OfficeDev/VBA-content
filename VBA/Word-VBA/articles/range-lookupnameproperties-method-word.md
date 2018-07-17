---
title: Range.LookupNameProperties Method (Word)
keywords: vbawd10.chm157155505
f1_keywords:
- vbawd10.chm157155505
ms.prod: word
api_name:
- Word.Range.LookupNameProperties
ms.assetid: a3a0facf-898a-d8c9-033a-b48416b53266
ms.date: 06/08/2017
---


# Range.LookupNameProperties Method (Word)

Looks up a name in the global address book list and displays the  **Properties** dialog box, which includes information about the specified name.


## Syntax

 _expression_ . **LookupNameProperties**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

If this method finds more than one match, it displays the  **Check Names** dialog box.


## Example

This example looks up the selected name in the address book and displays the  **Properties** dialog box for that person.


```
Selection.Range.LookupNameProperties
```


## See also


#### Concepts


[Range Object](range-object-word.md)

