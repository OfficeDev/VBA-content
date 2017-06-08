---
title: Document.SelectSingleNode Method (Word)
keywords: vbawd10.chm158007784
f1_keywords:
- vbawd10.chm158007784
ms.prod: word
api_name:
- Word.Document.SelectSingleNode
ms.assetid: 85f22e41-97e3-4413-c57e-26719155dc7d
ms.date: 06/08/2017
---


# Document.SelectSingleNode Method (Word)

Returns an  **XMLNode** object that represents the first node that matches the XPath parameter in the specified document.


## Syntax

 _expression_ . **SelectSingleNode**( **_XPath_** , **_PrefixMapping_** , **_FastSearchSkippingTextNodes_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|A valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.|
| _PrefixMapping_|Optional| **Variant**|Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.|
| _FastSearchSkippingTextNodes_|Optional| **Boolean**| **True** skips all text nodes while searching for the specified node. **False** includes text nodes in the search. Default value is **True** .|

### Return Value

XMLNode


## Remarks

Setting the FastSearchSkippingTextNodes parameter to  **False** diminishes performance because Microsoft Word searches all nodes in a document against the text contained in the node.


## Example

The following example returns the first title element found in the active document that is a child element of the book element.


```vb
Dim objElement As XMLNode 
Dim strElement As String 
Dim strPrefix As String 
 
strElement = "/x:catalog/x:book/x:title" 
strPrefix = "xmlns:x=""" &; ActiveDocument _ 
 .XMLSchemaReferences(1).NamespaceURI &; """" 
 
Set objElement = ActiveDocument _ 
 .SelectSingleNode(strElement, strPrefix)
```


## See also


#### Concepts


[Document Object](document-object-word.md)

