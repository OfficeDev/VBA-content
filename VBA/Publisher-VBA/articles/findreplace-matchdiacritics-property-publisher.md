---
title: FindReplace.MatchDiacritics Property (Publisher)
keywords: vbapb10.chm8323081
f1_keywords:
- vbapb10.chm8323081
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchDiacritics
ms.assetid: e23d01a1-9252-4077-c52f-87c53b5c0589
ms.date: 06/08/2017
---


# FindReplace.MatchDiacritics Property (Publisher)

Sets or returns a  **Boolean** representing whether or not a search operation will match diacritics. Read/write.


## Syntax

 _expression_. **MatchDiacritics**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Remarks

This property may not be available depending on the languages enabled on your operating system. The default value is  **False**.

Returns  **Access denied** if a proper language, such as Arabic, is not enabled.


## Example

This example finds the first occurrence of the word "gegenüber" in a German document. 


```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "gegenüber" 
 .MatchDiacritics = True 
 .Execute 
End With 

```


