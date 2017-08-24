---
title: FindReplace.MatchWidth Property (Publisher)
keywords: vbapb10.chm8323084
f1_keywords:
- vbapb10.chm8323084
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchWidth
ms.assetid: b9f89092-6ac0-bbf9-4bfd-d3cce2359b80
ms.date: 06/08/2017
---


# FindReplace.MatchWidth Property (Publisher)

Sets or returns a  **Boolean** representing whether or not a search operation will match the character width of the searched text. Read/Write.


## Syntax

 _expression_. **MatchWidth**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Remarks

This property may not be available depending on the language enabled on your operating system. The default value is  **False**.

Return "Access denied" if an East Asian language is not enabled.


## Example

The following example finds each occurance of the word "width" in the active document and applies bold formatting. The  **MatchWidth** property is set to **False** so that full or half width characters will both be found. For example, this search will apply bold formatting to the word "width" (half-width characters) and the word " w i d t h" (full-width characters).


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "width" 
 .MatchWidth = False 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```

The following example finds each occurance of the word "width" in the active document and applies bold formatting. The  **MatchWidth** property is set to **True** so that either full or half width characters will be found. For example, this search will apply bold formatting to "width". It will not apply formatting to the word "w i d t h".




```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "width" 
 .MatchWidth = True 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```


