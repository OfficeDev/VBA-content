---
title: ExchangeUser.GetFreeBusy Method (Outlook)
keywords: vbaol11.chm2075
f1_keywords:
- vbaol11.chm2075
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetFreeBusy
ms.assetid: 0dcd36af-e9d7-ca1e-334f-c540c46254f7
ms.date: 06/08/2017
---


# ExchangeUser.GetFreeBusy Method (Outlook)

Obtains a  **String** representing the availability of the **[ExchangeUser](exchangeuser-object-outlook.md)** for a period of 30 days from the start date, beginning at midnight of the date specified.


## Syntax

 _expression_ . **GetFreeBusy**( **_Start_** , **_MinPerChar_** , **_CompleteFormat_** )

 _expression_ A variable that represents an **ExchangeUser** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Date**|The date of availability, starting at midnight.|
| _MinPerChar_|Required| **Long**|Specifies the length of each time slot in minutes. Default is 30 minutes.|
| _CompleteFormat_|Optional| **Variant**|A value of  **True** indicates that a finer granularity of busy time is returned in the free/busy string. A value of **False** indicates that a distinction between only the available and busy times is returned.|

### Return Value

A  **String** that represents the availability of the Exchange user for a period of 30 days from the start date, beginning at midnight of the date specified. Each character in the **String** is a value indicating if the user is available (0), and optionally, whether a busy time is marked tentative (1), out of office (3), or other (2).


## Example

The following Visual Basic for Applications (VBA) example uses the  **GetFreeBusy** method to retrieve the free/busy information, with each time slot representing a 60 minute period, for the manager assigned to the current user. The example then uses that information to calculate the date and time on which the first free period occurs, and displays that information in the **Debug** window.


```vb
Sub GetManagerOpenInterval() 
 Dim oManager As ExchangeUser 
 Dim oCurrentUser As ExchangeUser 
 Dim FreeBusy As String 
 Dim BusySlot As Long 
 Dim DateBusySlot As Date 
 Dim i As Long 
 Const SlotLength = 60 
 'Get ExchangeUser for CurrentUser 
 If Application.Session.CurrentUser.AddressEntry.Type = "EX" Then 
 Set oCurrentUser = _ 
 Application.Session.CurrentUser.AddressEntry.GetExchangeUser 
 'Get Manager 
 Set oManager = oManager.GetExchangeUserManager 
 If oManager Is Nothing Then 
 Exit Sub 
 End If 
 FreeBusy = oManager.GetFreeBusy(Now, SlotLength) 
 For i = 1 To Len(FreeBusy) 
 If CLng(Mid(FreeBusy, i, 1)) = 0 Then 
 'get the number of minutes into the day for free interval 
 BusySlot = (i - 1) * SlotLength 
 'get an actual date/time 
 DateBusySlot = DateAdd("n", BusySlot, Date) 
 'To refine this function, substitute actual 
 'workdays and working hours in date/time comparison 
 If TimeValue(DateBusySlot) >= TimeValue(#8:00:00 AM#) And _ 
 TimeValue(DateBusySlot) <= TimeValue(#5:00:00 PM#) And _ 
 Not (Weekday(DateBusySlot) = vbSaturday Or _ 
 Weekday(DateBusySlot) = vbSunday) Then 
 Debug.Print oManager.name &; " first open interval:" &; _ 
 vbCrLf &; _ 
 Format$(DateBusySlot, "dddd, mmm d yyyy hh:mm AMPM") 
 Exit For 
 End If 
 End If 
 Next 
 End If 
End Sub
```


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

