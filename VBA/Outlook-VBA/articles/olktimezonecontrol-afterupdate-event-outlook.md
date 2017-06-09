---
title: OlkTimeZoneControl.AfterUpdate Event (Outlook)
keywords: vbaol11.chm1000528
f1_keywords:
- vbaol11.chm1000528
ms.prod: outlook
api_name:
- Outlook.OlkTimeZoneControl.AfterUpdate
ms.assetid: b34419cd-3df9-6855-032a-8ed7193a82fb
ms.date: 06/08/2017
---


# OlkTimeZoneControl.AfterUpdate Event (Outlook)

Occurs after the data in the control has been changed through the user interface.


## Syntax

 _expression_ . **AfterUpdate**

 _expression_ A variable that represents an **OlkTimeZoneControl** object.


## Remarks

 **[BeforeUpdate](olktimezonecontrol-beforeupdate-event-outlook.md)** and **AfterUpdate** can occur any time the data in the control is being saved to the item. The typical sequence of events involving **AfterUpdate** for this control is as follows:


1. User focuses on the control
    
2.  **BeforeUpdate** occurs
    
3. Control data is updated
    
4.  **AfterUpdate** occurs
    
5.  **[Exit](olktimezonecontrol-exit-event-outlook.md)** occurs: User moves focus away from control
    



## See also


#### Concepts


[OlkTimeZoneControl Object](olktimezonecontrol-object-outlook.md)

