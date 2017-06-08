---
title: Presentation.SendFaxOverInternet Method (PowerPoint)
keywords: vbapp10.chm583085
f1_keywords:
- vbapp10.chm583085
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SendFaxOverInternet
ms.assetid: 4470cafb-16f5-045b-1dab-8f8ead50ffe0
ms.date: 06/08/2017
---


# Presentation.SendFaxOverInternet Method (PowerPoint)

Sends a presentation as a fax to the specified recipients.


## Syntax

 _expression_. **SendFaxOverInternet**( **_Recipients_**, **_Subject_**, **_ShowMessage_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional|**Variant**|A  **String** that represents the fax numbers and e-mail addresses of the people to whom to send the fax. Separate multiple recipients with a semicolon.|
| _Subject_|Optional|**Variant**|A  **String** that represents the subject line for the faxed presentation.|
| _ShowMessage_|Optional|**Variant**|Whether to display the fax message before sending it.  **True** displays the fax message before sending it. **False** sends the fax without displaying the fax message.|

## Remarks

Using the  **SendFaxOverInternet** method requires that the fax service be enabled on a user's computer.

The format used for specifying fax numbers in the Recipients parameter is either  _recipientsfaxnumber_ @ _usersfaxprovider_ or _recipientsname_ @ _recipientsfaxnumber_. You can access the user's fax provider information by using the following registry path:

```text
HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax
```

Use the  `FaxAddress` key value under the above registry path to determine the format to use for a user.


## Example

The following example sends a fax to the fax service provider, who will fax the message to the recipient.


```vb
ActivePresentation.SendFaxOverInternet _
    "14255550101@consolidatedmessenger.com", _
    "For your review", True
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

