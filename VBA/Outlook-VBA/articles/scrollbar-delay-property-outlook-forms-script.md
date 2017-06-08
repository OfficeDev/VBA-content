---
title: ScrollBar.Delay Property (Outlook Forms Script)
keywords: olfm10.chm2001080
f1_keywords:
- olfm10.chm2001080
ms.prod: outlook
ms.assetid: 6eb415e2-961e-09df-2504-a1c65ebd3b32
ms.date: 06/08/2017
---


# ScrollBar.Delay Property (Outlook Forms Script)

Returns or sets a  **Long** that specifies the delay in milliseconds, between events on a **[ScrollBar](scrollbar-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Delay**

 _expression_A variable that represents a  **ScrollBar** object.


## Remarks

The  **Delay** property affects the amount of time between consecutive **SpinUp**,  **SpinDown**, and  **Change** events generated when the user clicks and holds down a button on a **ScrollBar**. The first event occurs immediately. The delay to the second occurrence of the event is five times the value of the specified  **Delay**. This initial lag makes it easy to generate a single event rather than a stream of events.

After the initial lag, the interval between events is the value specified for  **Delay**.

The default value of  **Delay** is 50 milliseconds. This means the object initiates the first event after 250 milliseconds (5 times the specified value) and initiates each subsequent event after 50 milliseconds.


