---
title: Application.AddAddress Method (Word)
keywords: vbawd10.chm158335297
f1_keywords:
- vbawd10.chm158335297
ms.prod: word
api_name:
- Word.Application.AddAddress
ms.assetid: 9114f213-9e43-a65c-7513-631820481967
ms.date: 06/08/2017
---


# Application.AddAddress Method (Word)

Adds an entry to the address book. Each entry has values for one or more tag IDs.

## Syntax

_expression_. **AddAddress** (**_TagID_**, **_Value_**)

_expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TagID_|Required|**String array**|The tag ID values for the new address entry. Each element in the array can contain one of the strings listed in the [following table](#tag-ids). Only the display name is required; the remaining entries are optional.|
| _Value_|Required|**String array**|The values for the new address entry. Each element corresponds to an element in the TagID array. For more information, see the [example](#example).|

<br/>

#### Tag IDs

|**Tag ID**|**Description**|
|:-----|:-----|
|PR_DISPLAY_NAME|Name displayed in the **Address Book** dialog box|
|PR_DISPLAY_NAME_PREFIX|Title (for example, "Ms." or "Dr.")|
|PR_GIVEN_NAME|First name|
|PR_SURNAME|Last name|
|PR_STREET_ADDRESS|Street address|
|PR_LOCALITY|City or locality|
|PR_STATE_OR_PROVINCE|State or province|
|PR_POSTAL_CODE|Postal code|
|PR_COUNTRY|Country/Region|
|PR_TITLE|Job title|
|PR_COMPANY_NAME|Company name|
|PR_DEPARTMENT_NAME|Department name within the company|
|PR_OFFICE_LOCATION|Office location|
|PR_PRIMARY_TELEPHONE_NUMBER|Primary telephone number|
|PR_PRIMARY_FAX_NUMBER|Primary fax number|
|PR_OFFICE_TELEPHONE_NUMBER|Office telephone number|
|PR_OFFICE2_TELEPHONE_NUMBER|Second office telephone number|
|PR_HOME_TELEPHONE_NUMBER|Home telephone number|
|PR_CELLULAR_TELEPHONE_NUMBER|Cellular telephone number|
|PR_BEEPER_TELEPHONE_NUMBER|Beeper telephone number|
|PR_COMMENT|Text included on the **Notes** tab for the address entry|
|PR_EMAIL_ADDRESS|Electronic mail address|
|PR_ADDRTYPE|Electronic mail address type|
|PR_OTHER_TELEPHONE_NUMBER|Alternate telephone number (other than home or office)|
|PR_BUSINESS_FAX_NUMBER|Business fax number|
|PR_HOME_FAX_NUMBER|Home fax number|
|PR_RADIO_TELEPHONE_NUMBER|Radio telephone number|
|PR_INITIALS|Initials|
|PR_LOCATION|Location, in the format buildingnumber/roomnumber (for example, 7/3007 represents room 3007 in building 7)|
|PR_CAR_TELEPHONE_NUMBER|Car telephone number|

<br/>

## Example

This example adds an entry to the address book.

```vb
Dim tagIDArray(0 To 3) As String 
Dim valueArray(0 To 3) As String 
 
tagIDArray(0) = "PR_DISPLAY_NAME" 
tagIDArray(1) = "PR_GIVEN_NAME" 
tagIDArray(2) = "PR_SURNAME" 
tagIDArray(3) = "PR_COMMENT" 
valueArray(0) = "Kim Buhler" 
valueArray(1) = "Kim" 
valueArray(2) = "Buhler" 
valueArray(3) = "This is a comment" 
 
Application.AddAddress TagID:=tagIDArray(), Value:=valueArray()
```


## See also

#### Concepts

- [Application Object](application-object-word.md)

