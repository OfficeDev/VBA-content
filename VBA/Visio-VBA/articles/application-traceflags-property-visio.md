---
title: Application.TraceFlags Property (Visio)
keywords: vis_sdr.chm10014590
f1_keywords:
- vis_sdr.chm10014590
ms.prod: visio
api_name:
- Visio.Application.TraceFlags
ms.assetid: aae7879a-7f21-f16e-cfcc-2520c70af7e7
ms.date: 06/08/2017
---


# Application.TraceFlags Property (Visio)

Gets or sets events logged during a Microsoft Visio instance. Read/write.


## Syntax

 _expression_ . **TraceFlags**

 _expression_ A variable that represents an **Application** object.


### Return Value

Long


## Remarks

The value of the  **TraceFlags** property can be a combination of the following values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visTraceEvents**|&;H1|Event occurrences|
| **visTraceAdvises**|&;H2|Outgoing advise calls|
| **visTraceAddonInvokes**|&;H4|Add-on invocations|
| **visTraceCallsToVBA**|&;H8|VBA invocations|
Setting the  **visTraceEvents** flag causes Visio to log most events as they happen and display them in the Immediate window. In most cases this occurs even if no external agent is listening or responding to the event. In a few cases, Visio knows there is no listener for an event and does not log those events. Visio also does not log idle events or advises. In addition, some events are specializations of other events and aren't recorded. For example, the **SelectionAdded** event is manufactured from distinct **ShapeAdded** events, so the Immediate window records the **ShapeAdded** events but not the **SelectionAdded** events.

Here is a string Visio might log when  **visTraceEvents** is set:




```
-event: 0x8040 /doc=1 /page=1 /shape=Sheet.1
```

The number after -event: is the code of the event that occurred. In this case 0x8040 is the code for the  **ShapeAdded** event. The text following the event code differs from event to event.

Setting the  **visTraceAdvises** flag writes a line to the Immediate window just before Visio calls an event handler procedure and another line just after the event handler returns. This includes event procedures in Microsoft Visual Basic for Applications (VBA) projects, for example, procedures in **ThisDocument** . Here is an example of what you might see:




```
>advise seq=4 event=0x8040 sink=0x40097598 
<advise seq=4 

```

These strings indicate the call to and return from an event handler. The sequence number also indicates this event was the fourth one fired by Visio. The code of the event is 0x8040 and the address of the interface Visio called is 0x40097598.

Setting the  **visTraceAddonInvokes** flag records when Visio invokes an EXE or VSL add-on, and when Visio regains control. Here is an example:




```
>invokeAO: SHOWARGS.EXE 
<invokeAO: completed 

```

Setting the  **visTraceAddonInvokes** flag also traces attempts to invoke add-ons that are not present. For example, if a cell's formula is =RunAddon("xxx") and there is no add-on named "xxx", the message "InvokeAO: Failed to map 'xxx' to known Add-on" is logged.

Setting the  **visTraceCallToVBA flag** writes a line to the Immediate window just before it makes a call to VBA other than a call to an event procedure (use **visTraceAdvises** to log calls to VBA event procedures) and another line just after VBA returns control to Visio. This flag traces macro invocations, calls to VBA procedures resulting from evaluation of cells that make use of RunAddon or CallThis operands, and calls resulting from selection of custom menu or toolbar items. Here is an example:




```
>invokeVBA: Module1.MyMacro 
<invokeVBA: completed 

```

A message doesn't appear in the Immediate window unless a document that has a VBA project is open. Visio queues a small number of messages to log when such a document opens. However, messages are lost if no document with a project is available for lengthy periods. Messages are also lost if VBA resets or if there are undismissed breakpoints.

Code in VBA projects can intersperse its messages with those logged by Visio by using standard  **Debug.Print** statements. Code in non-VBA projects can log messages to the Immediate window by using Document.VBProject.ExecuteLine("Debug.Print ""somestring""").

The  **TraceFlags** property is recorded in the **TraceFlags** entry of the **Application** section of the registry.


## Example

This VBA macro shows how to use the  **TraceFlags** property to log events, advises, add-on invocations, and Visual Basic invocations in the Immediate window.


```vb
 
Public Sub TraceFlags_Example() 
 
 Application.TraceFlags = visTraceEvents + visTraceAdvises + _ 
 visTraceAddonInvokes + visTraceCallsToVBA 
 
End Sub
```


