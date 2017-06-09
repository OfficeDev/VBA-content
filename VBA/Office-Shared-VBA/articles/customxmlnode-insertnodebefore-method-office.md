---
title: CustomXMLNode.InsertNodeBefore Method (Office)
keywords: vbaof11.chm294023
f1_keywords:
- vbaof11.chm294023
ms.prod: office
api_name:
- Office.CustomXMLNode.InsertNodeBefore
ms.assetid: b2805906-16b7-aebd-ccde-ded736a1b69b
ms.date: 06/08/2017
---


# CustomXMLNode.InsertNodeBefore Method (Office)

Inserts a new node just before the context node in the tree.


## Syntax

 _expression_. **InsertNodeBefore**( **_Name_**, **_NamespaceURI_**, **_NodeType_**, **_NodeValue_**, **_NextSibling_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|Represents the base name of the node to be added.|
| _NamespaceURI_|Optional|**String**|Represents the namespace of the element to be added. This parameter is required if adding nodes of type  **msoCustomXMLNodeElement** or **msoCustomXMLNodeAttribute**, otherwise it is ignored.|
| _NodeType_|Optional|**MsoCustomXMLNodeType**|Specifies the type of the node to be added. If the parameter is not specified, it is assumed to be a node of type  **msoCustomXMLNodeElement**.|
| _NodeValue_|Optional|**String**|Used to set the value of the node to be added for those nodes that allow text. If the node doesn't allow text, the parameter is ignored.|
| _NextSibling_|Optional|**CustomXMLNode**|Represents the context node.|

## Remarks

If the context node is not present when adding a node of type  **msoCustomXMLNodeElement**, **msoCustomXMLNodeComment**, or **msoCustomXMLNodeProcessingInstruction**, the new node is added to the last child of the context node. If the operation would result in an invalid tree structure, the insertion is not performed and an error message is displayed.


## Example

The following example adds a custom part and then finds a node in that part by using an XPath expression. The code then inserts a node before the found node.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
With ActiveDocument 
 
   ' Add a custom xml part. 
   .CustomXMLParts.Add "<invoice>" 
         
 
   ' Returns the first custom xml part with the given root namespace. 
   Set cxp1 = .CustomXMLParts("urn:invoice:namespace")              
  
   ' Get node using XPath.                              
   Set cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]")  
  
   ' Insert a node before the single node selected previously. 
    cxn.InsertNodeBefore("discount", "urn:invoice:namespace")   
              
 End With
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

