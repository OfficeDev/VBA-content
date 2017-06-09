---
title: FindReplace.MatchAlefHamza Property (Publisher)
keywords: vbapb10.chm8323079
f1_keywords:
- vbapb10.chm8323079
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchAlefHamza
ms.assetid: a8bdfbc3-13b5-e6a1-d86c-95e8f58ec263
ms.date: 06/08/2017
---


# FindReplace.MatchAlefHamza Property (Publisher)

Sets or returns a  **Boolean** representing whether or not a search operation will match alefs and hamzas. Read/write.


## Syntax

 _expression_. **MatchAlefHamza**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Remarks

This property may not be available depending on the language enabled on your operating system. The default value is  **False**.

Returns  **Access denied** if Arabic is not enabled.


## Example

This example finds the first occurrence of the word "" in an Arabic document matching alefs and hamzas.


```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "" 
 .MatchAlefHamza = True 
 .Execute 
End With 

```

This example follows from the previous one except that alef hamzas will not be matched. Therefore the words "" or "" will both be found because alefs and hamzas will be ignored.




```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "" 
 .MatchAlefHamza = False 
 .Execute 
End With 

```


