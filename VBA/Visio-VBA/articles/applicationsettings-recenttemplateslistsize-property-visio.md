---
title: ApplicationSettings.RecentTemplatesListSize Property (Visio)
keywords: vis_sdr.chm16262515
f1_keywords:
- vis_sdr.chm16262515
ms.prod: visio
api_name:
- Visio.ApplicationSettings.RecentTemplatesListSize
ms.assetid: a9b40755-31c9-a297-fe32-e9e0939d32fc
ms.date: 06/08/2017
---


# ApplicationSettings.RecentTemplatesListSize Property (Visio)

Determines the number of entries in the  **Recent Templates** list in the Microsoft Visio user interface. Read/write.


## Syntax

 _expression_ . **RecentTemplatesListSize**

 _expression_ A variable that represents an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

 **Long**


## Remarks

The value of the  **RecentTemplatesListSize** property corresponds to the setting of the **Show this number of Recent Templates** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box. To open the **Visio Options** dialog box, click the **File** tab, click **Options**, and then click  **Advanced**. The maximum number of recently used templates that can be displayed is 12.


## Example

The following Visual Basic for Applications (VBA) macro shows how to use the  **RecentTemplatesListSize** property to print the current size of the recently used template list in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **[Application](application-object-visio.md)** object.


```vb
Public Sub RecentTemplatesListSize_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngListSize As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngListSize = vsoApplicationSettings.RecentTemplatesListSize 
 
    Debug.Print lngListSize 
 
End Sub
```


