---
title: Characters.AddCustomFieldU Method (Visio)
keywords: vis_sdr.chm10251925
f1_keywords:
- vis_sdr.chm10251925
ms.prod: visio
api_name:
- Visio.Characters.AddCustomFieldU
ms.assetid: f1a5bc23-981d-0be7-92f3-d2ba640751a2
ms.date: 06/08/2017
---


# Characters.AddCustomFieldU Method (Visio)

Replaces the text represented by a  **Characters** object with a custom formula field that uses universal syntax.


## Syntax

 _expression_ . **AddCustomFieldU**( **_Formula_** , **_Format_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Formula_|Required| **String**|The formula of the new field.|
| _Format_|Required| **Integer**|The format of the new field.|

### Return Value

Nothing


## Remarks

Using the  **AddCustomFieldU** method is similar to clicking **Field** on the **Insert** tab and inserting a custom formula field in text. To add any other type of field (not custom), use the **AddField** method.

Valid field format constants are defined in the Visio type library in  **[VisFieldFormats](visfieldformats-enumeration-visio.md)** .


 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **AddCustomField** method to set a custom field that uses local syntax. Use the **AddCustomFieldU** method to set a custom field that uses universal syntax.


