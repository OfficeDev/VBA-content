---
title: Application.SaveAsTemplate Method (Access)
keywords: vbaac10.chm14524
f1_keywords:
- vbaac10.chm14524
ms.prod: access
api_name:
- Access.Application.SaveAsTemplate
ms.assetid: 3f796181-70c7-f372-92e9-0c2dbbc7262a
ms.date: 06/08/2017
---


# Application.SaveAsTemplate Method (Access)

Converts an existing Access database file to a database template (*.accdt) format file.


## Syntax

 _expression_. **SaveAsTemplate**( ** _Path_**, ** _Title_**, ** _IconPath_**, ** _CoreTable_**, ** _Category_**, ** _PreviewPath_**, ** _Description_**, ** _InstantiationForm_**, ** _ApplicationPart_**, ** _IncludeData_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The full path and file name of the database template to create.|
| _Title_|Required|**String**|The name of the database that is created when the instantiates the template.|
| _IconPath_|Required|**String**|An image file to be used as an icon for the database template.|
| _CoreTable_|Required|**String**|The table that contains the data that users would most want to create a relationship with when they instantiate the template. The Application Part argument must be set to  **True** if you set use this argument.|
| _Category_|Required|**String**|The template category under which the database template will appear on the  **Available Templates** page.|
| _PreviewPath_|Optional|**Variant**|An image file to be used as a preview for the database template on the  **Available Templates** page.|
| _Description_|Optional|**Variant**|A description to be displayed when the user selects the database template in the  **Available Templates** page.|
| _InstantiationForm_|Optional|**Variant**|Specifies the name of the form to be displayed when the template is instantiated.|
| _ApplicationPart_|Optional|**Variant**|Specifies whether the template will be displayed when the user clicks  **Application Parts** in the **Templates** group of the **Create** ribbon tab. Set to **True** to display the template when the user clicks **Application Parts**.|
| _IncludeData_|Optional|**Variant**|Specifies whether the table data is included in the template. Set to  **True** to include the table data.|

## Remarks

The  **SaveAsTemplate** method replaces the **SaveAsTemplate** method that was installed with the Microsoft Office Access 2007 Developer Extensions.

Templates created by the  **SaveAsTemplate** method cannot be used in Microsoft Office Access 2007 or earlier versions.


## See also


#### Concepts


[Application Object](application-object-access.md)

