---
title: CustomXMLNode.SelectNodes Method (Office)
keywords: vbaof11.chm294028
f1_keywords:
- vbaof11.chm294028
ms.prod: office
api_name:
- Office.CustomXMLNode.SelectNodes
ms.assetid: 443592af-a684-ee5e-98af-3e157f0f135e
ms.date: 06/08/2017
---


# CustomXMLNode.SelectNodes Method (Office)

Selects a collection of nodes matching an XPath expression. This method differs from the  **CustomXMLPart**. **SelectNodes** method in that the XPath expression will be evaluated starting with the 'expression' node as the context node.


## Syntax

 _expression_. **SelectNodes**( **_XPath_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains an XPath expression.|

### Return Value

CustomXMLNodes


## Example

The following example demonstrates adding a custom XML part, selecting a part matching a namespace URI, and then selecting nodes within that part that match an XPath expression.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Add a custom xml part. 
ActiveDocument.CustomXMLParts.Add "<supplier>" 
 
' Return the first custom xml part with the given namespace. 
Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")  
 
' Get all of the nodes matching an XPath expression. 
 Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

