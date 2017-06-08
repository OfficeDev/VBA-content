---
title: CustomXMLPart.NodeAfterInsert Event (Office)
keywords: vbaof11.chm296001
f1_keywords:
- vbaof11.chm296001
ms.prod: office
api_name:
- Office.CustomXMLPart.NodeAfterInsert
ms.assetid: 7ea1ce05-9992-608b-bac9-95f5d80ff586
ms.date: 06/08/2017
---


# CustomXMLPart.NodeAfterInsert Event (Office)

Occurs after a node is inserted in a  **CustomXMLPart** object.


## Syntax

 _expression_. **NodeAfterInsert**( **_NewNode_**, **_InUndoRedo_** )

 _expression_ An expression that returns a **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewNode_|Required|**CustomXMLNode**|Corresponds to the node just added to the  **CustomXMLPart** object. Note that this node may have children, if a subtree was just added to the document.|
| _InUndoRedo_|Required|**Boolean**|Returns  **TRUE** if the node was inserted as part of an Undo/Redo action by the user.|

## Example

The following example displays a message telling the user the results of inserting the node.


```
Sub CustomXMLParts_NodeAfterInsert(newNode As CustomXMLNode, boolInUndoRedo As Boolean) 
   MsgBox ("The node " &amp; newNode.BaseName &amp; " was just inserted.") 
End Sub
```


## See also


#### Concepts


[CustomXMLPart Object](customxmlpart-object-office.md)
#### Other resources


[CustomXMLPart Object Members](customxmlpart-members-office.md)

