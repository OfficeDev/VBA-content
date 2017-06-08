---
title: Report.PrtDevNames Property (Access)
keywords: vbaac10.chm13739
f1_keywords:
- vbaac10.chm13739
ms.prod: access
api_name:
- Access.Report.PrtDevNames
ms.assetid: 96a3437b-3655-5a87-9a1f-722116c82708
ms.date: 06/08/2017
---


# Report.PrtDevNames Property (Access)

You can use the  **PrtDevNames** property to set or return information about the printer selected in the **Print** dialog box for a form or report. Read/write **Variant**.


## Syntax

 _expression_. **PrtDevNames**

 _expression_ A variable that represents a **Report** object.


## Remarks

It is strongly recommended that you consult the Win32 Software Development Kit for complete documentation on the  **PrtDevMode**, **PrtDevNames**, and **PrtMip** properties.

The  **PrtDevNames** property is a variable-length structure that mirrors the DEVNAMES structure defined in the Win32 Software Development Kit.

The  **PrtDevNames** property uses the following members.



|**Member**|**Description**|
|:-----|:-----|
|DriverOffset|Specifies the offset from the beginning of the structure to a  **Null-** terminated string that specifies the file name (without an extension) of the device driver. This string is used to specify which printer is initially displayed in the **Print** dialog box.|
|DeviceOffset|Specifies the offset from the beginning of the structure to the  **Null-** terminated string that specifies the name of the device. This string can't be longer than 32 bytes (including the null character) and must be identical to the DeviceName member of the DEVMODE structure.|
|OutputOffset|Specifies the offset from the beginning of the structure to the  **Null-** terminated string that specifies the MS-DOS device name for the physical output medium (output port); for example, "LPT1:".|
|Default|Specifies whether the strings specified in the DEVNAMES structure identify the default printer. Before the  **Print** dialog box is displayed, if Default is set to 1 and all of the values in the DEVNAMES structure match the current default printer, the selected printer is set to the default printer. Default is set to 1 if the current default printer has been selected.|
Microsoft Access sets the  **PrtDevNames** property when you make selections in the Printer section of the **Print** dialog box. You can also set the property by using Visual Basic .

Microsoft Access uses the DEVNAMES structure to initialize the  **Print** dialog box. When the user chooses **OK** to close the dialog box, information about the selected printer is returned by the **PrtDevNames** property.


## See also


#### Concepts


[Report Object](report-object-access.md)

