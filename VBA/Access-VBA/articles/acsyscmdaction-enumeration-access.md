---
title: AcSysCmdAction Enumeration (Access)
keywords: vbaac10.chm10027
f1_keywords:
- vbaac10.chm10027
ms.prod: access
api_name:
- Access.AcSysCmdAction
ms.assetid: a2879d50-9845-40b0-9e51-a022340c664b
ms.date: 06/08/2017
---


# AcSysCmdAction Enumeration (Access)

Used with the  **SysCmd** method to specify an action to take.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**acSysCmdAccessDir**|9|Returns the name of the directory where Msaccess.exe is located.|
|**acSysCmdAccessVer**|7|Returns the version number of Microsoft Access.|
|**acSysCmdClearHelpTopic**|11||
|**acSysCmdClearStatus**|5|Provides information on the state of a database object.|
|**acSysCmdGetObjectState**|10|Returns the state of the specified database object. You must specify argument1 and argument2 when you use this action value.|
|**acSysCmdGetWorkgroupFile**|13|Returns the path to the workgroup file (System.mdw).|
|**acSysCmdIniFile**|8|Returns the name of the .ini file associated with Microsoft Access.|
|**acSysCmdInitMeter**|1|Initializes the progress meter. You must specify the argument1 and argument2 arguments when you use this action.|
|**acSysCmdProfile**|12|Returns the ** /profile** setting specified by the user when starting Microsoft Access from the command line.|
|**acSysCmdRemoveMeter**|3|Removes the progress meter.|
|**acSysCmdRuntime**|6|Returns  **True** (?1) if a run-time version of Microsoft Access is running.|
|**acSysCmdSetStatus**|4|Sets the status bar text to the text argument.|
|**acSysCmdUpdateMeter**|2|Updates the progress meter with the specified value. You must specify the text argument when you use this action.|

