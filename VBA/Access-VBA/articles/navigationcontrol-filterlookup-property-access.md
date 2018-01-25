---
title: NavigationControl.FilterLookup Property (Access)
keywords: vbaac10.chm11062
f1_keywords:
- vbaac10.chm11062
ms.prod: access
api_name:
- Access.NavigationControl.FilterLookup
ms.assetid: c368853c-6a1c-f104-2180-ebc889cf7e6d
ms.date: 06/08/2017
---


# NavigationControl.FilterLookup Property (Access)

You can use the  **FilterLookup** property to specify whether values appear in a boundtext box control when using the Filter By Form or Server Filter By Form window. Read/write **Byte**.


## Syntax

 _expression_. **FilterLookup**

 _expression_ A variable that represents a **NavigationControl** object.


## Remarks

The  **FilterLookup** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Never|0|The field values aren't displayed. You can specify whether the filtered records can contain null values.|
|Database Default|1|(Default) The field values are displayed according to the settings under  **Filter lookup options** on the **Current Database** tab of the **Access Options** dialog box, available by clicking the **Microsoft Office Button** ![File menu button](images/O12FileMenuButton_ZA10077102.gif) and the clicking **Access Options**.|
|Always|2|The field values are always displayed.|

You can also set the default for this property by setting a control's  **DefaultControl** property in Visual Basic.


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

