---
title: CustomXMLPart.SelectSingleNode Method (Office)
keywords: vbaof11.chm295013
f1_keywords:
- vbaof11.chm295013
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLPart.SelectSingleNode
ms.assetid: 2bd4c25b-d4e6-08db-b2ce-c74adf16336f
---


# CustomXMLPart.SelectSingleNode Method (Office)

Selects a single node within a custom XML part matching an XPath expression.


## Syntax

 _expression_. **SelectSingleNode**( ** _XPath_** )

 _expression_ An expression that returns a **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPath_|Required|**String**|Contains an XPath expression.|

### Return Value

CustomXMLNode


## Example

The following example demonstrates adding a custom XML part, selecting a part with a namespace URI, and then selecting a node within that part that matches an XPath expression. 


```vb
Dim cxp1 As CustomXMLPart
Dim cxn As CustomXMLNode

' Add a custom XML part.
ActiveDocument.CustomXMLParts.Add ( _
    "<suppliers>" &; _
    "<supplier ID='1'>Contoso</supplier>" &; _
    "<supplier ID='2'>Wingtip Toys</supplier>" &; _
    "</suppliers>")

' Return the last custom XML part added to the document.
Set cxp1 = ActiveDocument.CustomXMLParts(ActiveDocument.CustomXMLParts.Count)

' Get a node using XPath.
Set cxn = cxp1.SelectSingleNode("//supplier[@ID=1]")

' Display the node value 'Contoso'.
MsgBox cxn.NodeValue


```


## See also


#### Concepts


[CustomXMLPart Object](customxmlpart-object-office.md)

