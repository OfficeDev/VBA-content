---
title: Application.CustomizeIMEMode Method (Project)
keywords: vbapj.chm254
f1_keywords:
- vbapj.chm254
ms.prod: project-server
api_name:
- Project.Application.CustomizeIMEMode
ms.assetid: 1e6cae3d-7b06-327a-4db1-8b4416d703ee
ms.date: 06/08/2017
---


# Application.CustomizeIMEMode Method (Project)

Customizes which IME mode is used on a given field.


## Syntax

 _expression_. **CustomizeIMEMode**( ** _FieldID_**, ** _IMEMode_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldID_|Optional|**Long**|The field to customize. The default value is  **pjTaskName**. Can be one of the **[PjField](pjfield-enumeration-project.md)** constants|
| _IMEMode_|Optional|**Long**|Specifies the IME mode to use when the focus is on a table column. The default value is  **pjIMEModeNoControl**. Can be one of the **[PjIMEMode](pjimemode-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

The  **CustomizeIMEMode** method produces tangible results only if an East Asian version of Project is used.

Using the  **CustomizeIMEMode** method without specifying any arguments displays the **Customize IME Mode** dialog box.


