---
title: Application.OnDataChangeDelay Property (Visio)
keywords: vis_sdr.chm10013970
f1_keywords:
- vis_sdr.chm10013970
ms.prod: visio
api_name:
- Visio.Application.OnDataChangeDelay
ms.assetid: 14952e41-445a-77ff-30f7-e7aa6d8fcc32
ms.date: 06/08/2017
---


# Application.OnDataChangeDelay Property (Visio)

Gets or sets how long the Microsoft Visio instance waits before advising a container application that a Visio document being shown by the container has changed and should be redisplayed. Read/write.


## Syntax

 _expression_ . **OnDataChangeDelay**

 _expression_ An expression that returns a **Application** object.


### Return Value

Long


## Remarks

The  **OnDataChangeDelay** property affects only instances of Visio that are run from within an OLE container document.


- Setting the value of the  **OnDataChangeDelay** property to zero (0) causes Visio to send immediate advises to the container as data changes in open Visio documents.
    
- Setting the value of the  **OnDataChangeDelay** property to -1 causes Visio to use the interval specified in the **OLEUpdateDelay** entry in the registry. If the registry doesn't contain this setting, Visio defaults to using a value of 10,000 (milliseconds).
    
- Setting the  **OnDataChangeDelay** property to any value other than -1 or 0 overrides the registry setting and sets the delay between advises to the value of **OnDataChangeDelay** . If the **OnDataChangeDelay** property is not set or set to 1 and the **OLEUpdateDelay** setting is 0, Visio never sends advises to the container.
    

 **Note**  If you experience decreased performance when you add or move shapes that are contained in a Visio drawing that is hosted in the Visio Drawing Control, you can install a hot fix and then use the  **OnDataChangeDelay** property to overcome this problem. For more information, see the Microsoft Knowledge Base article[839079](http://support.microsoft.com/?scid=http%3a%2f%2fwww.support.microsoft.com%2fkb%2f839079%2f).


