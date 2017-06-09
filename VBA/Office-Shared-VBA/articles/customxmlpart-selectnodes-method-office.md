---
title: CustomXMLPart.SelectNodes Method (Office)
keywords: vbaof11.chm295012
f1_keywords:
- vbaof11.chm295012
ms.prod: office
api_name:
- Office.CustomXMLPart.SelectNodes
ms.assetid: c220c535-ac3f-cdba-5b1b-b608ed2eb8e4
ms.date: 06/08/2017
---


# CustomXMLPart.SelectNodes Method (Office)

Selects a collection of nodes from a custom XML part.


## Syntax

 _expression_. **SelectNodes**( **_XPath_** )

 _expression_ An expression that returns a **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains the XPath expression.|

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


[CustomXMLPart Object](customxmlpart-object-office.md)
#### Other resources


[CustomXMLPart Object Members](customxmlpart-members-office.md)

