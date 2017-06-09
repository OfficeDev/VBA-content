---
title: ApplicationSettings.RecentFilesListSize Property (Visio)
keywords: vis_sdr.chm16251680
f1_keywords:
- vis_sdr.chm16251680
ms.prod: visio
api_name:
- Visio.ApplicationSettings.RecentFilesListSize
ms.assetid: 8057f3d5-ccaf-28a2-9e70-1844f858d51d
ms.date: 06/08/2017
---


# ApplicationSettings.RecentFilesListSize Property (Visio)

Determines the number of entries in the  **Recent Documents** list in the Microsoft Visio user interface. Read/write.


## Syntax

 _expression_ . **RecentFilesListSize**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

Long


## Remarks

Setting the  **RecentFilesListSize** property is equivalent to setting the **Show this number of Recent Documents** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (click the **File** tab, click **Options**, and then click  **Advanced**). the maximum setting is 12.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **RecentFilesListSize** property to print the current size of the recently used file list in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub RecentFilesListSize_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngListSize As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngListSize = vsoApplicationSettings.RecentFilesListSize 
 
    Debug.Print lngListSize 
 
End Sub
```


