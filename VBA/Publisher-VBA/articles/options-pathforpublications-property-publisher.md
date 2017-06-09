---
title: Options.PathForPublications Property (Publisher)
keywords: vbapb10.chm1048597
f1_keywords:
- vbapb10.chm1048597
ms.prod: publisher
api_name:
- Publisher.Options.PathForPublications
ms.assetid: d33d5eab-eb52-b533-8968-31ddb5e12d99
ms.date: 06/08/2017
---


# Options.PathForPublications Property (Publisher)

Returns a  **String** that represents the default folder for publications. Read.


## Syntax

 _expression_. **PathForPublications**

 _expression_A variable that represents a  **Options** object.


### Return Value

String


## Example

This example returns the current default path for publications (corresponds to the default path setting on the  **General** tab in the **Options** dialog box, **Tools** menu).


```vb
Sub PubPath() 
 Dim strPubPath 
 strPubPath = Options.PathForPublications 
 MsgBox strPubPath 
End Sub
```


