---
title: CustomXMLPart.NodeAfterDelete Event (Office)
keywords: vbaof11.chm296002
f1_keywords:
- vbaof11.chm296002
ms.prod: office
api_name:
- Office.CustomXMLPart.NodeAfterDelete
ms.assetid: 430d2eed-afc3-8798-1478-2146351cefcc
ms.date: 06/08/2017
---


# CustomXMLPart.NodeAfterDelete Event (Office)

Occurs after a node is deleted in a  **CustomXMLPart** object.


## Syntax

 _expression_. **NodeAfterDelete**( **_OldNode_**, **_OldParentNode_**, **_OldNextSibling_**, **_InUndoRedo_** )

 _expression_ An expression that returns a **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OldNode_|Required|**CustomXMLNode**|Corresponds to the node which was just removed from the  **CustomXMLPart** object. Note that this node may have children, if a subtree is being removed from the document. Also, this node will be a "disconnected" node in that you can query down from the node, but you cannot query up the tree - the node appears to exist alone.|
| _OldParentNode_|Required|**CustomXMLNode**|Corresponds to the former parent node of OldNode.|
| _OldNextSibling_|Required|**CustomXMLNode**|Corresponds to the former next sibling of OldNode.|
| _InUndoRedo_|Required|**Boolean**|Returns  **TRUE** if the node was inserted as part of an Undo/Redo action by the user.|

## Example

The following example displays a message telling the user the results of deleting the node.


```
Sub CustomXMLParts_NodeAfterDelete(newNode As CustomXMLNode, boolInUndoRedo As Boolean) 
   MsgBox ("The node " &amp; newNode.BaseName &amp; " was just deleted.") 
End Sub
```


## See also


#### Concepts


[CustomXMLPart Object](customxmlpart-object-office.md)
#### Other resources


[CustomXMLPart Object Members](customxmlpart-members-office.md)

