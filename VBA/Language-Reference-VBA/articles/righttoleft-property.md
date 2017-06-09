---
title: RightToLeft Property
keywords: vblr6.chm1264146
f1_keywords:
- vblr6.chm1264146
ms.prod: office
api_name:
- Office.RightToLeft
ms.assetid: 0d3678c4-57c4-4c7c-aa2f-77ce1c339524
ms.date: 06/08/2017
---


# RightToLeft Property



Returns a Boolean value that indicates the text display direction and controls the visual appearance on a bidirectional system.
 **Remarks**
The settings for the  **RightToLeft** property are:


|**Setting**|**Description**|
|:-----|:-----|
|**True**|The control is running on a bidirectional platform, such as Arabic Windows95 or Hebrew Windows95, and text is running from right to left. The control should modify its behavior, such as putting vertical scroll bars at the left side of a text or list box, putting labels to the right of text boxes, and so forth.|
|**False**|The control should act as though it was running on a non-bidirectional platform, such as English Windows95, and text is running from left to right. If the container does not implement this ambient property,  **False** is the default value.|
The  **RightToLeft** property holds ambient information from the **Userform** that suggests behavior to controls contained within the **UserForm**. This property is a Boolean that indicates text display direction and controls visual appearance on bidirectional systems. The default is **False**.

