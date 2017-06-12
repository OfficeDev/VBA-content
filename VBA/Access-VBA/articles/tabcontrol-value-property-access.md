---
title: TabControl.Value Property (Access)
keywords: vbaac10.chm12071
f1_keywords:
- vbaac10.chm12071
ms.prod: access
api_name:
- Access.TabControl.Value
ms.assetid: 85849d32-3ef9-b959-fe07-026de226623e
ms.date: 06/08/2017
---


# TabControl.Value Property (Access)

Determines or specifies the selected  **[Page](page-object-access.md)** object. Read/write **Variant**.


## Syntax

 _expression_. **Value**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **Value** property of a tab control contains the index number of the current **Page** object. There is one **Page** object for each tab in a tab control. The first **Page** object always has an index number of 0, the second has an index number of 1, and so on.

The  **Value** property returns or sets a control's default property, which is the property that is assumed when you don't explicitly specify a property name.


 **Note**   The **Value** property is not the same as the **DefaultValue** property, which specifies the value that a property is assigned when a new record is created.


## See also


#### Concepts


[TabControl Object](tabcontrol-object-access.md)

