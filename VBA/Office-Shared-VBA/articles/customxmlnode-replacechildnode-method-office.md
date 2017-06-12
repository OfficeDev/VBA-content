---
title: CustomXMLNode.ReplaceChildNode Method (Office)
keywords: vbaof11.chm294026
f1_keywords:
- vbaof11.chm294026
ms.prod: office
api_name:
- Office.CustomXMLNode.ReplaceChildNode
ms.assetid: 72d571f4-8a54-b250-ce5d-22d595ef09f4
ms.date: 06/08/2017
---


# CustomXMLNode.ReplaceChildNode Method (Office)

Removes the specified child node (and its subtree) from the main tree, and replaces it with a different node in the same location.


## Syntax

 _expression_. **ReplaceChildNode**( **_OldNode_**, **_Name_**, **_NamespaceURI_**, **_NodeType_**, **_NodeValue_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OldNode_|Required|**CustomXMLNode**|Represents the child node to be replaced.|
| _Name_|Optional|**String**|Represents the base name of the element to be added.|
| _NamespaceURI_|Optional|**String**|Represents the namespace of the element to be added. This parameter is required if adding nodes of type  **msoCustomXMLNodeElement** or **msoCustomXMLNodeAttribute**, otherwise it is ignored.|
| _NodeType_|Optional|**MsoCustomXMLNodeType**|Specifies the type of node to add. If the parameter is not specified, it is assumed to be of type  **msoCustomXMLNodeElement**.|
| _NodeValue_|Optional|**String**|Used to set the value of the node to be added for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.|

## Remarks

If the  _OldNode_ parameter is not a child of the context node or if the operation would result in an invalid tree structure, the replacement is not performed and an error message is displayed. In addition, in a case where the node to be added already exists, the replacement is not performed and an error message is displayed.


## Example

The following example selects a custom part and then a node in that part. The code then replaces a child of that node with another node.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
With ActiveDocument 
 
   ' Return the first custom xml part with the given root namespace. 
   Set cxp1 = .CustomXMLParts("urn:invoice:namespace")     '  
                              
   Set cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")  
 
   ' Replace a child node. 
    cxn.ReplaceChildNode(cxn.SelectSingleNode("//discount", "rebate")   
        
End With
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

