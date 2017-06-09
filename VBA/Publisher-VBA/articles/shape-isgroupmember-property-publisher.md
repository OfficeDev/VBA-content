---
title: Shape.IsGroupMember Property (Publisher)
keywords: vbapb10.chm2228337
f1_keywords:
- vbapb10.chm2228337
ms.prod: publisher
api_name:
- Publisher.Shape.IsGroupMember
ms.assetid: bbd9b662-b47d-d5cf-6858-e208c44f88a0
ms.date: 06/08/2017
---


# Shape.IsGroupMember Property (Publisher)

Returns  **True** if the specified shape is a member of a group, **False** otherwise. Read-only **Boolean**.


## Syntax

 _expression_. **IsGroupMember**

 _expression_A variable that represents an  **Shape** object.


### Return Value

Boolean


## Remarks

The object returned by the  **ParentGroupShape** property can be used to determine the parent shape for the group.


## Example

The following statement can be used to return a  **True** value if the first shape of the active publication is a group member.


```
blnGrouped = Application.ActiveDocument.MasterPages _ 
 .Item.Shapes(1).IsGroupMember
```


