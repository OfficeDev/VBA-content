---
title: Application.DontResetInsertionPointProperties Property (Word)
keywords: vbawd10.chm158335456
f1_keywords:
- vbawd10.chm158335456
ms.prod: word
api_name:
- Word.Application.DontResetInsertionPointProperties
ms.assetid: 3e6dfd03-9ab9-43c2-378c-0d97c69e14b3
ms.date: 06/08/2017
---


# Application.DontResetInsertionPointProperties Property (Word)

Returns or sets a  **Boolean** that represents whether Microsoft Word maintains the formatting properties of the text at that position of the Insertion Point after running other code. Read/write.


## Syntax

 _expression_ . **DontResetInsertionPointProperties**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

In some cases, Word loses the formatting at the Insertion Point after running other Microsoft Visual Basic for Applications (VBA) code. When this happens, it can cause difficulty for users who rely on a screen reader application. They lose the formatting when their assistive application performs what seems like unrelated tasks. This property prevents Word from losing or changing the formatting that has been applied to the text at the position of the Insertion Point when other code runs that contains properties or methods in the Word object model.


 **Important**  Do not use this property unless you specifically need it to make a solution function correctly.


## See also


#### Concepts


[Application Object](application-object-word.md)

