---
title: DrawingControl.HostID Property (Visio)
keywords: vis_sdr.chm52000
f1_keywords:
- vis_sdr.chm52000
ms.prod: visio
api_name:
- Visio.HostID
ms.assetid: ecc77cb3-04c8-6a31-0d40-d03dddb6bf20
ms.date: 06/08/2017
---


# DrawingControl.HostID Property (Visio)

String representation of the GUID of the host application of the Microsoft Visio Drawing Control. Read/write.


## Syntax

 _expression_ . **HostID**

 _expression_ A variable that represents a **DrawingControl** object.


### Return Value

String


## Remarks

Setting this property is optional. Set  **HostID** at design time, for example in the **Properties** window in your Microsoft Visual Studio project.

Setting the  **HostID** property creates (or modifies) a subkey at the following location in the registry:

 ```text
 HKEY_CURRENTUSER\Software\Microsoft\Office\14.0\VisioHosts\
 ```

If you set  **HostID** , use a unique string that identifies your program, preferably a unique GUID, although any string less than 128 characters or less in length that contains no backslash ("\") or forward slash ("/") character is valid. A GUID should be no more than 40 characters. Write your Setup program so that when it uninstalls your program, it deletes the registry key and all its subkeys.

All instances of the Visio Drawing Control in your program share the same  **HostID** value. If you set **HostID** multiple times for multiple instances of the Visio Drawing Control within the same application, later values will overwrite earlier ones in the registry.

 **HostID** data is global to all instances of the Drawing Control in a given Win32 process. If you change any of the registry settings associated with any instance of the control at the **HostID** location, that setting is changed for all instances.

If you leave  **HostID** set to the default, an empty string (""), your application will share application settings, for example, those you can set in the **Visio Options** dialog box (click the **File** tab, and then click **Options**), with Visio itself. If you set a  **HostID** value, your application will have its own persistent registry-based values for these settings. In addition, you thereby leave Visio settings unchanged. If you have more than one application that contains a Drawing Control and you want each to have its own application settings, assign each a different **HostID** .


 **Caution**   Modifying the Microsoft Windows registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


