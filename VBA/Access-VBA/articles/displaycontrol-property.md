---
title: DisplayControl Property
keywords: vbaac10.chm4334
f1_keywords:
- vbaac10.chm4334
ms.prod: access
api_name:
- Access.DisplayControl
ms.assetid: 24deeaeb-22dc-b8fe-6c39-e9a2a4d12e2b
ms.date: 06/08/2017
---


# DisplayControl Property

  

**Applies to:** Access 2013 | Access 2016

You can use the  **DisplayControl** property in table Design view to specify the default control you want to use for displaying a field.


## Setting

You can set the  **DisplayControl** property in the table's property sheet in table Design view by clicking the **Lookup** tab in the **Field Properties** section.

This property contains a drop-down list of the available controls for the selected field. For fields with a Text or Number data type, this property can be set to Text Box, List Box, or Combo Box. For fields with a Yes/No data type, this property can be set to Check Box, Text Box, or Combo Box.


## Remarks

When you select a control for this property, any additional properties needed to configure the control are also displayed on the  **Lookup** tab.

Setting this property and any related control type properties will affect the field display in both Datasheet view and Form view. The field is displayed by using the control and control property settings set in table Design view. If a field had its  **DisplayControl** property set in table Design view and you drag it from the field list in form Design view, Microsoft Access copies the appropriate properties to the control's property sheet.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

