---
title: Recipient.FreeBusy Method (Outlook)
keywords: vbaol11.chm2357
f1_keywords:
- vbaol11.chm2357
ms.prod: outlook
api_name:
- Outlook.Recipient.FreeBusy
ms.assetid: eeb831bc-c369-10f1-fb0b-08a8105c48e6
ms.date: 06/08/2017
---


# Recipient.FreeBusy Method (Outlook)

Returns free/busy information for the recipient.


## Syntax

 _expression_ . **FreeBusy**( **_Start_** , **_MinPerChar_** , **_CompleteFormat_** )

 _expression_ A variable that represents a[Recipient](recipient-object-outlook.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Date**|The start date for the returned period of free/busy information.|
| _MinPerChar_|Required| **Long**|The number of minutes per character represented in the returned free/busy string.|
| _CompleteFormat_|Optional| **Variant**| **True** if the returned string should contain not only free/busy information, but also values for each character according to the **[OlBusyStatus](olbusystatus-enumeration-outlook.md)** constants.|

### Return Value

A  **String** value that represents the free/busy information.


## Remarks

 The default is to return a string representing one month of free/busy information compatible with the Microsoft Schedule+ Automation format (that is, the string contains one character for each _MinPerChar_ minute, up to one month of information from the specified _Start_ date).

If the optional argument  _CompleteFormat_ is omitted or **False** , then "free" is indicated by the character 0 and all other states by the character 1.

If  _CompleteFormat_ is **True** , then the same length string is returned as defined above, but the characters now correspond to the[OlBusyStatus](olbusystatus-enumeration-outlook.md) constants.


## Example

This Visual Basic for Applications (VBA) example uses the  **FreeBusy** method to return a string of free/busy information with one character for each day. This example allows for the possibility that the free/busy information for this recipient is not accessible. To run this example, you need to replace 'Nate Sun' with a valid recipient name.


```vb
Public Sub GetFreeBusyInfo() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim myRecipient As Outlook.Recipient 
 Dim myFBInfo As String 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 Set myRecipient = myNameSpace.CreateRecipient("Nate Sun") 
 On Error GoTo ErrorHandler 
 myFBInfo = myRecipient.FreeBusy(#11/11/2003#, 60 * 24) 
 MsgBox myFBInfo 
 Exit Sub 
ErrorHandler: 
 MsgBox "Cannot access the information. " 
End Sub
```

This VBA example returns a string of free/busy information with one character for each hour (complete format).




```vb
Set myRecipient = myNameSpace.CreateRecipient("Nate Sun") 
myFBInfo = myRecipient.FreeBusy(#8/1/03#, 60, True)
```


## See also


#### Concepts


[Recipient Object](recipient-object-outlook.md)

