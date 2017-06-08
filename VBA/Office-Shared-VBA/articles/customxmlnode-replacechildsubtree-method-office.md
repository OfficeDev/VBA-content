---
title: CustomXMLNode.ReplaceChildSubtree Method (Office)
keywords: vbaof11.chm294027
f1_keywords:
- vbaof11.chm294027
ms.prod: office
api_name:
- Office.CustomXMLNode.ReplaceChildSubtree
ms.assetid: 955ec2ab-c6c9-242c-5e05-3ff03b00b120
ms.date: 06/08/2017
---


# CustomXMLNode.ReplaceChildSubtree Method (Office)

Removes the specified node (and its subtree) from the main tree, and replaces it with a different subtree in the same location.


## Syntax

 _expression_. **ReplaceChildSubtree**( **_XML_**, **_OldNode_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XML_|Required|**String**|Represents the subtree to be added.|
| _OldNode_|Required|**CustomXMLNode**|Represents the child node to be replaced.|

## Remarks

If the operation would result in an invalid tree structure, the replacement operation is not performed and an error message is displayed. 


## Example

The following example selects a custom part and then a node in that part. The code then replaces a child subtree of that node with another subtree.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
With ActiveDocument 
 
    ' Return the first custom xml part with the given root namespace. 
    Set cxp1 = .CustomXMLParts("urn:invoice:namespace")     '  
         
    ' Get node using XPath expression.                              
    Set cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")  
 
    ' Replace one subtree and its children with another. 
   cxn.ReplaceChildSubtree("<rebates><rebate>0.10</rebate></rebates>", "//discounts")   
                 
 End With
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

