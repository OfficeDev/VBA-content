---
title: ViewCtl.Folder Property (Outlook View Control)
ms.prod: outlook
ms.assetid: c48e2d86-c284-9a25-0c11-00f6e24094c7
ms.date: 06/08/2017
---


# ViewCtl.Folder Property (Outlook View Control)

Returns or sets a  **String** that represents the path of the folder displayed by the control. Read/write.


## Syntax

 _expression_. **Folder**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

For security reasons, the return value includes the name of the root message store of the folder only if the folder is not in the user's personal mailbox. 

If neither the  **[Namespace](viewctl-namespace-property-outlook-view-control.md)** property nor the **Folder** property is set and the control is contained in a Microsoft Outlook folder home page, the control displays the current folder. If the **Namespace** property is set to "MAPI" and the **Folder** property is not set, the control displays the user's **Inbox**.

In addition to accepting a  **string** that represents a valid folder path, you can also set the **Folder** property to the EntryID of the folder that you want to display in the View Control.


