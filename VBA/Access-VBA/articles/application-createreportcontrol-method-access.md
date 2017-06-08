---
title: Application.CreateReportControl Method (Access)
keywords: vbaac10.chm12623
f1_keywords:
- vbaac10.chm12623
ms.prod: access
api_name:
- Access.Application.CreateReportControl
ms.assetid: 4b970377-450b-9909-f5c3-cb7f8445139f
ms.date: 06/08/2017
---


# Application.CreateReportControl Method (Access)

The  **CreateReportControl** method creates a control on a specified open report. For more information, see the **[CreateControl](application-createcontrol-method-access.md)** method.


## Syntax

 _expression_. **CreateReportControl**( ** _ReportName_**, ** _ControlType_**, ** _Section_**, ** _Parent_**, ** _ColumnName_**, ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReportName_|Required|**String**|The name of the open report on which you want to create the control.|
| _ControlType_|Required|**AcControlType**|An  **[AcControlType](accontroltype-enumeration-access.md)** constant that represents the type of control you want to create.|
| _Section_|Optional|**AcSection**|An  **[AcSection](acsection-enumeration-access.md)** constant that identifying the section that will contain the new control.|
| _Parent_|Optional|**Variant**|A string expression identifying the name of the parent control of an attached control. For controls that have no parent control, use a zero-length string for this argument, or omit it.|
| _ColumnName_|Optional|**Variant**| The name of the field to which the control will be bound, if it is to be a data-bound control.|
| _Left, Top_|Optional|**Variant**|The coordinates for the upper-left corner of the control in twips.|
| _Width, Height_|Optional|**Variant**|The width and height of the control in twips.|

### Return Value

Control


## See also


#### Concepts


[Application Object](application-object-access.md)

