---
title: GetFormat Method
keywords: fm20.chm2012310
f1_keywords:
- fm20.chm2012310
ms.prod: office
api_name:
- Office.GetFormat
ms.assetid: 4d056545-08c6-ef03-2980-1db42b01e6c9
ms.date: 06/08/2017
---


# GetFormat Method



Returns an integer value indicating whether a specific [format](glossary-vba.md) is on the **DataObject**.
 **Syntax**
 _Boolean_ = _object_. **GetFormat(**_format_**)**
The  **GetFormat** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _format_|Required. An integer or string specifying a specific format that might exist in the  **DataObject**. If the specified format exists in the **DataObject**, **GetFormat** returns **True**.|
 **Settings**
The settings for  _format_ are:


|**Value**|**Description**|
|:-----|:-----|
|1|Text format.|
|A string or any integer other than 1|A user-defined  **DataObject** format passed to the **DataObject** from **SetText**.|
 **Remarks**
The  **GetFormat** method searches for a format in the current list of formats on the **DataObject**. If the format is on the **DataObject**, **GetFormat** returns **True**; if not, **GetFormat** returns **False**.
The  **DataObject** currently supports only text formats.

