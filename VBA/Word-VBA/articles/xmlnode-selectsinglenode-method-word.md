---
title: XMLNode.SelectSingleNode Method (Word)
keywords: vbawd10.chm37748754
f1_keywords:
- vbawd10.chm37748754
ms.prod: word
api_name:
- Word.XMLNode.SelectSingleNode
ms.assetid: c831dba1-90f7-0af7-9e44-8f62a54de0fe
ms.date: 06/08/2017
---


# XMLNode.SelectSingleNode Method (Word)

Returns an  **XMLNode** object that represents the first child element that matches the XPath parameter within the specified XML element. .


## Syntax

 _expression_ . **SelectSingleNode**( **_XPath_** , **_PrefixMapping_** , **_FastSearchSkippingTextNodes_** )

 _expression_ An expression that returns an **[XMLNode](xmlnode-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|Specifies a valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.|
| _PrefixMapping_|Optional| **String**|Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.|
| _FastSearchSkippingTextNodes_|Optional| **Boolean**| **True** skips all text nodes while searching for the specified node. **False** includes text nodes in the search. Default value is **False** .|

### Return Value

XMLNode


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

