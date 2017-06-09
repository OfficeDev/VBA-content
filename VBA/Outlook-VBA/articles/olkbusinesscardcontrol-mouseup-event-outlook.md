---
title: OlkBusinessCardControl.MouseUp Event (Outlook)
keywords: vbaol11.chm1000332
f1_keywords:
- vbaol11.chm1000332
ms.prod: outlook
api_name:
- Outlook.OlkBusinessCardControl.MouseUp
ms.assetid: 877e7ab8-6b5e-f06a-db9a-9e9079110e9a
ms.date: 06/08/2017
---


# OlkBusinessCardControl.MouseUp Event (Outlook)

Occurs after the user releases a mouse button that has been pressed on the control.


## Syntax

 _expression_ . **MouseUp**( **_Button_** , **_Shift_** , **_X_** , **_Y_** )

 _expression_ A variable that represents an **OlkBusinessCardControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Integer**|An  **[OlMouseButton](olmousebutton-enumeration-outlook.md)** constant that specifies which button on the mouse has been pressed.|
| _Shift_|Required| **Integer**|A bitwise-OR mask of constants in the  **[OlShiftState](olshiftstate-enumeration-outlook.md)** enumeration that specifies whether the **SHIFT**,  **CTRL**, or  **ALT** keys have been pressed.|
| _X_|Required| **[OLE_XPOS_CONTAINER]**|Identifies the location of the mouse cursor on the X-axis relative to the form.|
| _Y_|Required| **[OLE_YPOS_CONTAINER]**|Identifies the location of the mouse cursor on the Y-axis relative to the form.|

## See also


#### Concepts


[OlkBusinessCardControl Object](olkbusinesscardcontrol-object-outlook.md)

