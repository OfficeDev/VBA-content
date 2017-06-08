---
title: ApplicationSettings.DeleteConnectorsEnabled Property (Visio)
keywords: vis_sdr.chm16262510
f1_keywords:
- vis_sdr.chm16262510
ms.prod: visio
api_name:
- Visio.ApplicationSettings.DeleteConnectorsEnabled
ms.assetid: adb52279-5837-08be-ce73-231656ef7640
ms.date: 06/08/2017
---


# ApplicationSettings.DeleteConnectorsEnabled Property (Visio)

Determines whether connectors are deleted when a shape to which they are connected is deleted. Read/write.


## Syntax

 _expression_ . **DeleteConnectorsEnabled**

 _expression_ A variable that represents an **[ApplicationSettings](applicationsettings-object-visio.md)** object.


### Return Value

Boolean


## Remarks

The setting of the  **DeleteConnectorsEnabled** property also determines whether two connectors reconnect when a shape to which they are both connected is deleted.

Connectors that contain text are not deleted when shapes to which they are connected are deleted, even if  **DeleteConnectorsEnabled** is **True** .

The setting of the  **DeleteConnectorsEnabled** property corresponds to the setting of the **Delete connectors when deleting shapes** check box under **Editing options** on the **Advanced** tab of the **Visio Options** dialog box. To open the **Visio Options** dialog box, click the **File** tab, click **Options**, and then click  **Advanced**.


