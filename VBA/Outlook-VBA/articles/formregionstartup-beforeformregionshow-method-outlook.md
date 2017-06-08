---
title: FormRegionStartup.BeforeFormRegionShow Method (Outlook)
keywords: vbaol11.chm2947
f1_keywords:
- vbaol11.chm2947
ms.prod: outlook
api_name:
- Outlook.FormRegionStartup.BeforeFormRegionShow
ms.assetid: c93c2f6a-511f-15cd-eca2-4eb35af9939a
ms.date: 06/08/2017
---


# FormRegionStartup.BeforeFormRegionShow Method (Outlook)

Allows an add-in to update the user interface of a form region before it is displayed. 


## Syntax

 _expression_ . **BeforeFormRegionShow**( **_FormRegion_** )

 _expression_ A variable that represents a **FormRegionStartup** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormRegion_|Required| **[FormRegion](formregion-object-outlook.md)**|The  **FormRegion** object representing the form region that is to be displayed.|

## Remarks

This method lets Outlook pass the  **FormRegion** object to the add-in, and allows an add-in to update the user interface of the form region before it is displayed so that, for instance, the text of labels can be changed or irrelevant content can be suppressed. It is called after the controls are instantiated and the layout is calculated, but before the form region is made visible.

When implementing this method, keep in mind that the item obtained from the  _FormRegion_ parameter (that is, the **[Item](formregion-item-property-outlook.md)** property of the **FormRegion** object) is read-only.

For examples of add-ins in C# and Visual Basic .NET that implement  **FormRegionStartup** , see code sample downloads on MSDN.


## See also


#### Concepts


[FormRegionStartup Interface](formregionstartup-object-outlook.md)

