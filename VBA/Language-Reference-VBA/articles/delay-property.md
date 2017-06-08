---
title: Delay Property
keywords: fm20.chm5225031
f1_keywords:
- fm20.chm5225031
ms.prod: office
api_name:
- Office.Delay
ms.assetid: 12d76300-bd1c-4b65-ca8e-b9c63e19100f
ms.date: 06/08/2017
---


# Delay Property



Specifies the delay for the SpinUp, SpinDown, and Change events on a  **SpinButton** or **ScrollBar**.
 **Syntax**
 _object_. **Delay** [= _Long_ ]
The  **Delay** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Long_|Optional. The delay, in milliseconds, between events.|
 **Remarks**
The  **Delay** property affects the amount of time between consecutive SpinUp, SpinDown, and Change events generated when the user clicks and holds down a button on a **SpinButton** or **ScrollBar**. The first event occurs immediately. The delay to the second occurrence of the event is five times the value of the specified **Delay**. This initial lag makes it easy to generate a single event rather than a stream of events.
After the initial lag, the interval between events is the value specified for  **Delay**.
The default value of  **Delay** is 50 milliseconds. This means the object initiates the first event after 250 milliseconds (5 times the specified value) and initiates each subsequent event after 50 milliseconds.

