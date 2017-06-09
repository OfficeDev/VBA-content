---
title: DataObject Object
keywords: fm20.chm2000510
f1_keywords:
- fm20.chm2000510
ms.prod: office
api_name:
- Office.DataObject
ms.assetid: 96ad2ab2-3e9b-2d7e-9502-a881e5dd8354
ms.date: 06/08/2017
---


# DataObject Object



A holding area for formatted text data used in transfer operations. Also holds a list of [formats](glossary-vba.md) corresponding to the pieces of text stored in the **DataObject**.
 **Remarks**
A  **DataObject** can contain one piece of text for the Clipboard text format, and one piece of text for each additional text format, such as custom and user-defined formats.
A  **DataObject** is distinct from the Clipboard. A **DataObject** supports commands that involve the Clipboard and drag-and-drop actions for text. When you start an operation involving the Clipboard (such as **GetText** ) or a drag-and-drop operation, the data involved in that operation is moved to a **DataObject**.
The  **DataObject** works like the Clipboard. If you copy a text string to a **DataObject**, the **DataObject** stores the text string. If you copy a second string of the same format to the **DataObject**, the **DataObject** discards the first text string and stores a copy of the second string. It stores one piece of text of a specified format and keeps the text from the most recent operation.

