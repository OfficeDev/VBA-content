---
title: MsoFeatureInstall Enumeration (Office)
ms.prod: office
api_name:
- Office.MsoFeatureInstall
ms.assetid: 25256738-d169-5c00-1d5d-eb8019811976
ms.date: 06/08/2017
---


# MsoFeatureInstall Enumeration (Office)

Specifies how the application handles calls to methods and properties that require features not yet installed.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
|**msoFeatureInstallNone**|0|Generates a generic automation error at run time when uninstalled features are called.|
|**msoFeatureInstallOnDemand**|1|Prompts the user to install new features.|
|**msoFeatureInstallOnDemandWithUI**|2|Displays a progress meter during installation; does not prompt the user to install new features.|

