---
title: CommandBars.GetImageMso Method (Office)
keywords: vbaof11.chm2025
f1_keywords:
- vbaof11.chm2025
ms.prod: office
api_name:
- Office.CommandBars.GetImageMso
ms.assetid: 36261e2b-9cbf-b0b6-5892-63bbb2f93959
ms.date: 06/08/2017
---


# CommandBars.GetImageMso Method (Office)

Returns an  **IPictureDisp** object of the control image identified by the **idMso** parameter scaled to the dimensions specified by width and height.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **GetImageMso**( **_idMso_**, **_Width_**, **_Height_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|
| _Width_|Required|**Integer**|The width of the image.|
| _Height_|Required|**Integer**|The height of the image.|

### Return Value

IPictureDisp


## Remarks

The  **Width** and **Height** parameters must be between 16 and 128.


## Example

The following sample returns a 32x32 version of the  **Paste** icon as an **IPictureDisp** object.


```
Application.CommandBars.GetImageMso("Paste", 32, 32)
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

