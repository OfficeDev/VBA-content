---
title: CustomXMLNode.AppendChildNode Method (Office)
keywords: vbaof11.chm294019
f1_keywords:
- vbaof11.chm294019
ms.prod: office
api_name:
- Office.CustomXMLNode.AppendChildNode
ms.assetid: 3fbe1c76-b60a-e365-4988-4a94a52e1fe0
ms.date: 06/08/2017
---


# CustomXMLNode.AppendChildNode Method (Office)

Appends a single node as the last child under the context element node in the tree. 


## Syntax

 _expression_. **AppendChildNode**( **_Name_**, **_NamespaceURI_**, **_NodeType_**, **_NodeValue_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|Represents the base name of the element to be appended.|
| _NamespaceURI_|Optional|**String**|Represents the namespace of the element to be appended. This parameter is required to append nodes of type  **msoCustomXMLNodeElement** or **msoCustomXMLNodeAttribute**, otherwise it is ignored.|
| _NodeType_|Optional|**MsoCustomXMLNodeType**|Specifies the type of node to append. If the parameter is not specified, it is assumed to be of type  **msoCustomXMLNodeElement**.|
| _NodeValue_|Optional|**String**|Used to set the value of the appended node for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.|

## Remarks

If the context node is any type other than  **msoXMLNodeElement**, or if the operation would result in an invalid tree structure, the append is not performed and an error message is displayed.


## Example

The following example demonstrates appending a  **CustomXMLNode** object to another node.


```
Sub AppendNode() 
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
 
        ' Add and populate a custom xml part 
        set cxp1 = .CustomXMLParts.Add "<invoice />" 
         
        ' Add a node 
        cxp1.AddNode "/invoice", "upccode", "urn:invoice:namespace" 
                        
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")  
 
        ' Append a child node to the single node selected previously. 
        cxn.AppendChildNode("discount", "urn:invoice:namespace", "string", "0.10")          
                         
    End With 
     
End Sub
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

