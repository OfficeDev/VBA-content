---
title: Workbook.SendFaxOverInternet Method (Excel)
keywords: vbaxl10.chm199223
f1_keywords:
- vbaxl10.chm199223
ms.prod: excel
api_name:
- Excel.Workbook.SendFaxOverInternet
ms.assetid: e7d91ac4-90d2-7555-af96-dc28736da769
ms.date: 06/08/2017
---


# Workbook.SendFaxOverInternet Method (Excel)

Sends a worksheet as a fax to the specfied recipients.


## Syntax

 _expression_ . **SendFaxOverInternet**( **_Recipients_** , **_Subject_** , **_ShowMessage_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional| **Variant**|A  **String** that represents the fax numbers and e-mail addresses of the people to whom the fax will be sent. Separate multiple recipients with a semicolon.|
| _Subject_|Optional| **Variant**|A  **String** that represents the subject line for the faxed document.|
| _ShowMessage_|Optional| **Variant**| **True** displays the fax message before sending it. **False** sends the fax without displaying the fax message.|

## Remarks

Using the  **SendFaxOverInternet** method requires that the fax service is enabled on a user's computer.

The format used for specifying fax numbers in the  _Recipients_ parameter is either "<recipientsfaxnumber>@<usersfaxprovider>" or "<recipientsname>@<recipientsfaxnumber>". You can access the user's fax provider information using the following registry path: `HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Common\Services\Fax`

Use the value of the FaxAddress key at this registry path to determine the format to use for a recipient.


## Example

The following example sends a fax to the fax service provider, which will then fax the message to the recipient.


```vb
ActiveWorkbook.SendFaxOverInternet _ 
 "14255550101@consolidatedmessenger.com", _ 
 "For your review", True
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

