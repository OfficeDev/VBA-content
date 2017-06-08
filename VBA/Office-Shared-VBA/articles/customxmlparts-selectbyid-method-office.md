---
title: CustomXMLParts.SelectByID Method (Office)
keywords: vbaof11.chm298005
f1_keywords:
- vbaof11.chm298005
ms.prod: office
api_name:
- Office.CustomXMLParts.SelectByID
ms.assetid: e9c0d3a1-c625-bb86-b4ca-6916d4a8a6b0
ms.date: 06/08/2017
---


# CustomXMLParts.SelectByID Method (Office)

Selects a custom XML part matching a GUID. 


## Syntax

 _expression_. **SelectByID**( **_Id_** )

 _expression_ An expression that returns a **CustomXMLParts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**String**|Contains the GUID for the custom XML part. |

### Return Value

CustomXMLPart


## Remarks

If a custom XML part with this ID does not exist, the method returns  **Nothing**.


## Example

The following example retrieves a custom XML part matching the GUID and then searches for a node in that part that matches an XPath expression.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Returns a custom xml part by its ID. 
 Set cxp1 = ActiveDocument.CustomXMLParts.SelectByID("F9168C5E-CEB2-4faa-B6BF-329BF39FA1E4")         
 
' Get the node matching the XPath expression.                              
Set cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")
```


## See also


#### Concepts


[CustomXMLParts Object](customxmlparts-object-office.md)
#### Other resources


[CustomXMLParts Object Members](customxmlparts-members-office.md)

