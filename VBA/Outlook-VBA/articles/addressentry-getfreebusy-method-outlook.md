---
title: AddressEntry.GetFreeBusy Method (Outlook)
keywords: vbaol11.chm2052
f1_keywords:
- vbaol11.chm2052
ms.prod: outlook
api_name:
- Outlook.AddressEntry.GetFreeBusy
ms.assetid: 8f3c7cbe-a4b5-ef5c-d7d3-1b38273f6f59
ms.date: 06/08/2017
---


# AddressEntry.GetFreeBusy Method (Outlook)

Returns a  **String** value that represents the availability of the individual user for a period of 30 days from the start date, beginning at midnight of the date specified.


## Syntax

 _expression_ . **GetFreeBusy**( **_Start_** , **_MinPerChar_** , **_CompleteFormat_** )

 _expression_ An expression that returns an **AddressEntry** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Date**|Specifies the date.|
| _MinPerChar_|Required| **Long**|Specifies the length of each time slot in minutes. The default value is 30.|
| _CompleteFormat_|Optional| **Variant**|Specifies a  **Boolean** value that represents the level of information returned for each time slot. The default value is **False** .|

### Return Value

A String value that represents the availability of the user for the specified period. The string value contains one character for each time slot within the specified period.


## Remarks


 **Note**  If an address entry represents a distribution list, the status of its individual members cannot be returned to you with the  **GetFreeBusy** method. A meeting request should be sent only to single messaging users. You can determine if a messaging user is a distribution list by determining if its **[DisplayType](addressentry-displaytype-property-outlook.md)** property is **olDistList** or **olPrivateDistList** .

The contents of the string returned by this method are determined by the  _CompleteFormat_ parameter.

If  _CompleteFormat_ is set to **False** , the default value, the string returned by this method contains one of the following characters for each time slot:



| **Character**| **Description**|
|0|The time slot represents a free period.|
|1|The time slot represents a tentative, out of office (OOF), or busy period.|
If  _CompleteFormat_ is set to **True** , the string returned by this method contains one of the following characters for each time slot:



| **Character**| **Description**|
|0|The time slot represents a free period.|
|1|The time slot represents a tentative period.|
|2|The time slot represents a busy period.|
|3|The time slot represents an out of office (OOF) period.|

## See also


#### Concepts


[AddressEntry Object](addressentry-object-outlook.md)

