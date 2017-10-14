---
title: CustomXMLNode.InsertSubtreeBefore Method (Office)
keywords: vbaof11.chm294024
f1_keywords:
- vbaof11.chm294024
ms.prod: office
api_name:
- Office.CustomXMLNode.InsertSubtreeBefore
ms.assetid: 5d9e9303-e427-a092-3960-eee90a53970d
ms.date: 06/08/2017
---


# CustomXMLNode.InsertSubtreeBefore Method (Office)

Inserts the specified subtree into the location just before the context node. 


## Syntax

 _expression_. **InsertSubtreeBefore**( **_XML_**, **_NextSibling_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XML_|Required|**String**|Represents the subtree to be added. |
| _NextSibling_|Optional|**CustomXMLNode**|Specifies the context node.|

## Remarks

If the  _NextSibling_ parameter is not a child of the context node or if the operation would result in an invalid tree structure, the insertion is not performed and an error message is displayed.


## Example

The following example adds a custom part and then finds a node in that part by using an XPath expression. The code then inserts a node after the found node.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
With ActiveDocument 
 
   ' Add a custom xml part. 
   .CustomXMLParts.Add "<invoice>"         
 
   ' Returns the first custom xml part with the given root namespace. 
   Set cxp1 = .CustomXMLParts("urn:invoice:namespace")              
  
   ' Get nodes using XPath.                              
   Set cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]")  
  
   ' Insert a node before the single node selected previously. 
    cxn.InsertNodeAfter("discount", "urn:invoice:namespace")   
              
 End With
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

