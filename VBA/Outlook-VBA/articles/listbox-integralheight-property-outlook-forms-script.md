---
title: ListBox.IntegralHeight Property (Outlook Forms Script)
keywords: olfm10.chm2001320
f1_keywords:
- olfm10.chm2001320
ms.prod: outlook
ms.assetid: b8574796-ec7a-c61a-4e87-cebb90220c5c
ms.date: 06/08/2017
---


# ListBox.IntegralHeight Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a **[ListBox](listbox-object-outlook-forms-script.md)** displays full lines of text in a list or partial lines. Read/write.


## Syntax

 _expression_. **IntegralHeight**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

 **True** indicates that the list resizes itself to display only complete items (default). **False** indicates that the list does not resize itself even if the item is too tall to display completely.

The  **IntegralHeight** property relates to the height of the list, just as the **AutoSize** property relates to the width of the list.

If  **IntegralHeight** is **True**, the list box automatically resizes when necessary to show full rows. If  **False**, the list remains a fixed size; if items are taller than the available space in the list, the entire item is not shown.


