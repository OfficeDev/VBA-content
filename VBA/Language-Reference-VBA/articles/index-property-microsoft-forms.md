---
title: Index Property (Microsoft Forms)
keywords: fm20.chm5225044
f1_keywords:
- fm20.chm5225044
ms.prod: office
ms.assetid: 304f42ff-5a38-0e84-8f9f-40e75d7fc2b2
ms.date: 06/08/2017
---


# Index Property (Microsoft Forms)



The position of a  **Tab** object within a **Tabs** collection or a **Page** object in a **Pages** collection.
 **Syntax**
 _object_. **Index** [= _Integer_ ]
The  **Index** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Integer_|Optional. The index of the currently selected  **Tab** object.|
 **Remarks**
The  **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of **Pages** in a **MultiPage** or **Tabs** on a **TabStrip**. The index value for the first page or tab is zero, the index value of the second page or tab is one, and so on.
In a  **MultiPage**, **Index** refers to a **Page** as well as the page's **Tab**. In a **TabStrip**, **Index** refers to the tab only.

