---
title: DoCmd.Quit Method (Access)
keywords: vbaac10.chm4167
f1_keywords:
- vbaac10.chm4167
ms.prod: access
api_name:
- Access.DoCmd.Quit
ms.assetid: 2644084a-fd24-6271-7679-46c5f1b206d5
ms.date: 06/08/2017
---


# DoCmd.Quit Method (Access)

The  **Quit** method quits Microsoft Access. You can select one of several options for saving a database object before quitting.


## Syntax

 _expression_. **Quit**( ** _Options_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Options_|Optional|**AcQuitOption**|An  **[AcQuitOption](acquitoption-enumeration-access.md)** constant that indicates the action to take when quitting Access. The default value is **acQuitSaveAll**.|

## Remarks

The  **Quit** method of the **DoCmd** object was added to provide backward compatibility for running the Quit action in Visual Basic code in Microsoft Access 95. It's recommended that you use the existing **Quit** method of the **[Application](application-object-access.md)** object instead.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

