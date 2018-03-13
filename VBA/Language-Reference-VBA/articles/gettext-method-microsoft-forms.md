---
title: GetText Method (Microsoft Forms)
keywords: fm20.chm2012320
f1_keywords:
- fm20.chm2012320
ms.prod: office
ms.assetid: 7d714405-4d3e-23e3-cedb-8a6a7fd07269
ms.date: 06/08/2017
---


# GetText Method (Microsoft Forms)



Retrieves a text string from the  **DataObject** using the specified[format](glossary-vba.md).
 **Syntax**
 _String_ = _object_. **GetText(** [ _format_ ] **)**
The  **GetText** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                       |
|:----------------------|:-------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                                     |
| <em>format</em>       | Optional. A string or integer specifying the format of the data to retrieve from the  <strong>DataObject</strong>. |

 **Settings**
The settings for  _format_ are:


| <strong>Value</strong>               | <strong>Description</strong>                                                                                                |
|:-------------------------------------|:----------------------------------------------------------------------------------------------------------------------------|
| 1                                    | Text format.                                                                                                                |
| A string or any integer other than 1 | A user-defined  <strong>DataObject</strong> format passed to the <strong>DataObject</strong> from <strong>SetText</strong>. |

 **Remarks**
The  **DataObject** supports multiple formats, but only supports one data item of each format. For example, the **DataObject** might include one text item and one item in a custom format; but cannot include two text items.
If no format is specified, the  **GetText** method requests information in the Text format from the **DataObject**.

