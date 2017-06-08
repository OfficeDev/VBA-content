---
title: CustomXMLParts.SelectByNamespace Method (Office)
keywords: vbaof11.chm298006
f1_keywords:
- vbaof11.chm298006
ms.prod: office
api_name:
- Office.CustomXMLParts.SelectByNamespace
ms.assetid: 39dcce9c-4354-0211-c2cf-393917bf6aef
ms.date: 06/08/2017
---


# CustomXMLParts.SelectByNamespace Method (Office)

Selects the collection of custom XML parts whose namespace matches the search criteria. 


## Syntax

 _expression_. **SelectByNamespace**( **_NamespaceURI_** )

 _expression_ An expression that returns a **CustomXMLParts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NamespaceURI_|Required|**String**|Contains a namespace URI.|

### Return Value

CustomXMLParts


## Remarks

If no custom XML parts with this namespace exist, the method returns an empty  **CustomXMLParts** collection object.


## Example

The following example selects all of the custom XML parts matching the namespace and then selects a node from those parts that match an XPath expression.


```
Dim cxp1 As CustomXMLParts 
Dim cxn As CustomXMLNode 
 
' Returns all of the custom xml parts with the given namespace. 
 Set cxp1 = ActiveDocument.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")    
 
' Get the node matching the XPath expression.                              
Set cxn = cxp1(1).SelectSingleNode("//*[@supplierID = 1]") 

```


## See also


#### Concepts


[CustomXMLParts Object](customxmlparts-object-office.md)
#### Other resources


[CustomXMLParts Object Members](customxmlparts-members-office.md)

