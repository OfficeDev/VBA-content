---
title: Source.Field Property (Word)
keywords: vbawd10.chm140836968
f1_keywords:
- vbawd10.chm140836968
ms.prod: word
api_name:
- Word.Source.Field
ms.assetid: fd6689d4-a042-4ca2-fddd-d048fe8c3a93
ms.date: 06/08/2017
---


# Source.Field Property (Word)

Returns a  **String** that represents the value of a field in a bibliography source. Read-only.


## Syntax

 _expression_ . **Field**( **_Name_** )

 _expression_ An expression that returns a **Source** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the field for which to retrieve the value.|

## Remarks

The name of the field corresponds to the name of the corresponding XML element in the resulting XML for a bibliography source. You can use the  **[XML](source-xml-property-word.md)** property to return the XML for a bibliography source. For more information, see[Working with Bibliographies](http://msdn.microsoft.com/library/ce05a0bd-bacd-16e1-0ab0-793a47a15da5%28Office.15%29.aspx).


## See also


#### Concepts


[Source Object](source-object-word.md)

