---
title: TableOfAuthorities.Passim Property (Word)
keywords: vbawd10.chm152109057
f1_keywords:
- vbawd10.chm152109057
ms.prod: word
api_name:
- Word.TableOfAuthorities.Passim
ms.assetid: 5df50485-69c7-ff9e-710c-7cdfdaaaeada
ms.date: 06/08/2017
---


# TableOfAuthorities.Passim Property (Word)

 **True** if five or more page references to the same authority are replaced with "Passim." Read/write **Boolean** .


## Syntax

 _expression_ . **Passim**

 _expression_ An expression that returns a **[TableOfAuthorities](tableofauthorities-object-word.md)** object.


## Remarks

Corresponds to the \p switch for a Table of Authorities (TOA) field.


## Example

This example formats the first table of authorities in Brief.doc to use page references instead of "Passim."


```vb
Documents("Brief.doc").TablesOfAuthorities(1).Passim = False
```

This example formats the tables of authorities in the active document to replace each instance of five or more page references for the same entry with "Passim."




```vb
For Each myTOA In ActiveDocument.TablesOfAuthorities 
 myToa.Passim = True 
Next myTOA
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

