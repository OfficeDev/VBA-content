---
title: CustomXMLNode.AppendChildSubtree Method (Office)
keywords: vbaof11.chm294020
f1_keywords:
- vbaof11.chm294020
ms.prod: office
api_name:
- Office.CustomXMLNode.AppendChildSubtree
ms.assetid: 67899ba9-7e5a-e40e-2e33-b02ff1fff4b4
ms.date: 06/08/2017
---


# CustomXMLNode.AppendChildSubtree Method (Office)

Adds a subtree as the last child under the context element node in the tree.


## Syntax

 _expression_. **AppendChildSubtree**( **_XML_** )

 _expression_ An expression that returns a **CustomXMLNode** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XML_|Required|**String**|Represents the subtree to add.|

## Remarks

If the context node is any type other than  **msoXMLNodeElement**, the append operation is not performed and an error message is displayed. If the CustomXMLNode is being validated against a schema and if the operation would result in an invalid tree structure, the append operation is not performed and an error message is displayed.


## Example

The following example demonstrates appending a node to an existing node.


```
Sub ShowCustomXmlParts() 
    Dim cxp1 As CustomXMLPart 
    Dim cxn As CustomXMLNode 
 
    With ActiveDocument 
 
        ' Add and populate a custom xml part 
        set cxp1 = .CustomXMLParts.Add "<invoice />" 
         
        ' Get nodes using XPath.                              
        Set cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")  
  
        ' Append a child subtree to the single node selected previously. 
        cxn.AppendChildSubtree("<discounts><discount>0.10</discount></discounts>")          
         
    End With 
     
End Sub
```


## See also


#### Concepts


[CustomXMLNode Object](customxmlnode-object-office.md)
#### Other resources


[CustomXMLNode Object Members](customxmlnode-members-office.md)

