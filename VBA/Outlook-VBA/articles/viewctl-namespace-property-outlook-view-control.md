---
title: ViewCtl.Namespace Property (Outlook View Control)
ms.prod: outlook
ms.assetid: 97cb1ea1-2e27-afc9-7756-b609dc9cc69e
ms.date: 06/08/2017
---


# ViewCtl.Namespace Property (Outlook View Control)

Returns or sets a  **String** that represents the namespace property of the control. Read/write.


## Syntax

 _expression_. **Namespace**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

If neither the  **Namespace** nor the **[Folder](viewctl-folder-property-outlook-view-control.md)** properties are set and the control is contained in a Microsoft Outlook folder home page, the control displays the current folder. If the **Namespace** property is set to "MAPI" and the **Folder** property is not set, the control displays the user's **Inbox**.

The namespace represents an abstract root object for any data source.


