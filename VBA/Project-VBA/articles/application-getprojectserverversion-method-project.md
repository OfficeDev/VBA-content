---
title: Application.GetProjectServerVersion Method (Project)
keywords: vbapj.chm131223
f1_keywords:
- vbapj.chm131223
ms.prod: project-server
api_name:
- Project.Application.GetProjectServerVersion
ms.assetid: f41cb738-3a30-f555-9d10-78343fae0ddb
ms.date: 06/08/2017
---


# Application.GetProjectServerVersion Method (Project)

This method checks the version of the Project Server for the active project. The method can also be used to check whether a particular server URL points to a valid and functioning Project Server.


## Syntax

 _expression_. **GetProjectServerVersion**( ** _ServerURL_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ServerURL_|Required|**String**|A string representing the URL of the Project Server whose version needs to be checked.|

### Return Value

PjServerVersionInfo


## Remarks

If the ServerURL argument does not point to a valid and functioning Project Server, the method returns a trappable error (error code 1004).


## Example

The following sample returns an XML stream representing the following settings from Project Server:  **ProjectServerSettingsRequest**, **AdminDefaultTrackingMethod**, **AdminTrackingLocked**, **ProjectIDInProjectServer**, **ProjectManagerHasTransactions**, **ProjectManagerHasTransactionsForCurrentProject**, **TimePeriodGranularity**, and **GroupsForCurrentProjectManager**.


```vb
Sub mpsVersion() 
 URL = ActiveProject.ServerURL 
 If Application.GetProjectServerVersion(URL) = pjServerVersionInfo_P10 Then 
 ActiveProject.MakeServerURLTrusted 
 xmlStream = Application.GetProjectServerSettings( _ 
 RequestXML:="<ProjectServerSettingsRequest>" _ 
 &; "<AdminDefaultTrackingMethod /><AdminTrackingLocked />" _ 
 &; "<ProjectIDInProjectServer />" _ 
 &; "<ProjectManagerHasTransactions />" _ 
 &; "<ProjectManagerHasTransactionsForCurrentProject />" _ 
 &; "<TimePeriodGranularity /><GroupsForCurrentProjectManager />" _ 
 &; "</ProjectServerSettingsRequest>") 
 MsgBox xmlStream 
 Else 
 MsgBox "This macro returns information from Project " _ 
 &; "Server. Please choose 'Collaborate using Project " _ 
 &; "Server' and specify a valid Project Server URL " _ 
 &; "for this project in Collaboration Options (Collaborate menu)." 
 Exit Sub 
 End If 
End Sub
```


