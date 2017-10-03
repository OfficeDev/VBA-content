---
title: TextBox.PostalAddress Property (Access)
keywords: vbaac10.chm11050
f1_keywords:
- vbaac10.chm11050
ms.prod: access
api_name:
- Access.TextBox.PostalAddress
ms.assetid: 04fb29c5-909c-a0b8-a4aa-7701abc07037
ms.date: 06/08/2017
---


# TextBox.PostalAddress Property (Access)

You can use the  **PostalAddress Property** property to specify or determine the postal code and the Customer Barcode data corresponding to the address information displayed in a specified field/textbox. The PostalAddress Property wizard enables the setting of these properties. Read/write **String**.


## Syntax

 _expression_. **PostalAddress**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

For processing the conversion, correct settings are necessary for all properties of field/textbox that will contain postal code, address, Customer Barcode data.

For settings, use section 1 to 3, delimiting with semicolon (;).

<br/>

**Setting for field/textbox for Postal code**

Specifies the type of postal code for the field/textbox.

|**Section**|**Description**|
|:-----|:-----|
|1|Specifies field/textbox for Prefecture names|
|2|Specifies field/textbox for City/Ward/County|
|3|Specifies field/textbox for Street/Town/Village|

<br/>

**Setting for field/textbox for address**

Specifies that the field/textbox contains a postal code or Customer Barcode data.

|**Section**|**Description**|
|:-----|:-----|
|1|Specifies field/textbox for postal code|
|2|Specifies field/textbox for Customer Barcode data|

**Note** Two semicolons are required at the end of the value. 

<br/>

**Setting for field/textbox for Customer Barcode data**

Specifies the type of Customer Barcode data in the field/textbox. This setting is the same as the field/textbox for postal code.

|**Section**|**Description**|
|:-----|:-----|
|1|Specifies field/textbox for Prefecture names|
|2|Specifies field/textbox for City/Ward/County|
|3|Specifies field/textbox for Street/Town/Village|

<br/>

The postal code consists of 3 address items: Prefecture, City/Ward/County, Street/Town/Village. Sections in the **PostalAddress Property** property of field/textbox for a postal code can be omitted. The following table shows how to omit sections from the property setting.

|**Property settings**|**Address input in field/textbox**|
|:-----|:-----|
|Address1|Address2 \| Address3|
|Address1|Prefecture+City/Ward/County+Street/Town/Village|
|Address1;|Prefecture|
|;Address1|City/Ward/County+Street/Town/Village|
|;Address1;|City/Ward/County|
|;;Address1|Street/Town/Village|
|Address1;Address2|Prefecture \| City/Ward/County+Street/Town/Village|
|Address1;Address1|Prefecture+City/Ward/County+Street/Town/Village|
|Address1;Address2;|Prefecture \| City/Ward/County|
|Address1;Address1;|Prefecture+City/Ward/County|
|;Address1;Address2|City/Ward/County \| Street/Town/Village|
|;Address1;Address1|City/Ward/County+Street/Town/Village|
|Address1;Address2;Address3|Prefecture \| City/Ward/County \| Street/Town/Village|
|Address1;Address2;Address2|Prefecture \| City/Ward/County+Street/Town/Village|
|Address1;Address1;Address2|Prefecture+City/Ward/County \| Street/Town/Village|
|Address1;Address1;Address1|Prefecture+City/Ward/County+Street/Town/Village|


The postal code converter program has been developed and licensed by Advanced Giken Corporation for Microsoft Corporation. 


## See also

#### Concepts

- [TextBox Object](textbox-object-access.md)

