---
title: User-Defined Data Type
keywords: vblr6.chm1009052
f1_keywords:
- vblr6.chm1009052
ms.prod: office
ms.assetid: 89ef52c6-f928-d43e-ef5d-8b6b3b5a3bce
ms.date: 06/08/2017
---


# User-Defined Data Type

Any [data type](vbe-glossary.md) you define using the **Type** statement. User-defined data types can contain one or more elements of a data type, an [array](vbe-glossary.md), or a previously defined user-defined type. For example:


```vb
Type MyType 
 MyName As String ' String variable stores a name. 
 MyBirthDate As Date ' Date variable stores a birthdate. 
 MySex As Integer ' Integer variable stores sex (0 for 
End Type ' female, 1 for male). 

```


