---
title: Document.SelectNodes Method (Word)
keywords: vbawd10.chm158007785
f1_keywords:
- vbawd10.chm158007785
ms.prod: word
api_name:
- Word.Document.SelectNodes
ms.assetid: b913720e-0f22-c626-6003-61a8dfb87f00
ms.date: 06/08/2017
---


# Document.SelectNodes Method (Word)

Returns an  **XMLNodes** collection that represents all the nodes that match the XPath parameter in the order in which they appear in the document or range.


## Syntax

 _expression_ . **SelectNodes**( **_XPath_** , **_PrefixMapping_** , **_FastSearchSkippingTextNodes_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required| **String**|A valid XPath string. For more information on XPath, see the XPath reference documentation on the Microsoft Developer Network (MSDN) Web site.|
| _PrefixMapping_|Optional| **Variant**|Provides the prefix in the schema against which to perform the search. Use the PrefixMapping parameter if your XPath parameter uses names to search for elements.|
| _FastSearchSkippingTextNodes_|Optional| **Boolean**| **True** skips all text nodes while searching for the specified node. **False** includes text nodes in the search. Default value is **False** .|

### Return Value

XMLNodes


## Remarks

Setting the FastSearchSkippingTextNodes parameter to  **True** diminishes performance, because Microsoft Word searches all nodes in a document against the text contained in the node.


## Example

The following example returns a collection of all book elements in the active document.


```vb
Dim objElements As XMLNodes 
Dim strElement As String 
Dim strPrefix As String 
 
strElement = "/x:catalog/x:book" 
strPrefix = "xmlns:x=""" &; ActiveDocument _ 
 .XMLSchemaReferences(1).NamespaceURI &; """" 
 
Set objElements = ActiveDocument _ 
 .SelectNodes(strElement, strPrefix)
```


## See also


#### Concepts


[Document Object](document-object-word.md)

