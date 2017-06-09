---
title: Broadcast.Capabilities Property (Word)
keywords: vbawd10.chm36438019
f1_keywords:
- vbawd10.chm36438019
ms.prod: word
ms.assetid: 86388adc-95c3-3c06-dbfe-a0455e93c90f
ms.date: 06/08/2017
---


# Broadcast.Capabilities Property (Word)

Returns a  **Long** that represents the capabilities of the specified broadcast. Read-only.


## Syntax

 _expression_ . **Capabilities**

 _expression_ A variable that represents a **Broadcast** object.


## Remarks

The  **Capabilities** property can return the following[MSOBroadcastCapabilities](http://msdn.microsoft.com/library/445ff0f7-fcb1-d65a-b055-189c268e2076%28Office.15%29.aspx) values:



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **MSOBroadcastCapFileSizeLimited**|1|File size limited|
| **MSOBroadcastCapSupportsMeetingNotes**|2|Supports meeting notes|
| **MSOBroadcastCapSupportsUpdateDoc**|4|Supports document update|
The values returned correspond to either Office or Microsoft Office 2010 broadcast presentation service capabilities.


## Property value

 **INT32**


## See also


#### Other resources


[Broadcast Object](broadcast-object-word.md)


