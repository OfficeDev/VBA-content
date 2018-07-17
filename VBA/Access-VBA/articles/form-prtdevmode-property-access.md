---
title: Form.PrtDevMode Property (Access)
keywords: vbaac10.chm13418
f1_keywords:
- vbaac10.chm13418
ms.prod: access
api_name:
- Access.Form.PrtDevMode
ms.assetid: a20a2dd9-4e5a-6fb7-63ba-8394e654057f
ms.date: 06/08/2017
---


# Form.PrtDevMode Property (Access)

You can use the  **PrtDevMode** property to set or return printing device mode information specified for a form or report in the **Print** dialog box. Read/write **Variant**.


## Syntax

 _expression_. **PrtDevMode**

 _expression_ A variable that represents a **Form** object.


## Remarks

It is strongly recommended that you consult the Win32 Software Development Kit for complete documentation on the  **PrtDevMode**, **PrtDevNames**, and **PrtMip** properties.

The  **PrtDevMode** property setting is a 94-byte structure that mirrors the DEVMODE structure defined in the Win32 Software Development Kit. For complete information on the **PrtDevMode** property members, consult the Win32 Software Development Kit.

The  **PrtDevMode** property uses the following members.



|**Member**|**Description**|
|:-----|:-----|
|DeviceName|A string with a maximum of 32 bytes that specifies the name of the device the driver supports ? for example, "HP LaserJet IIISi" if the Hewlett-Packard LaserJet IIISi is the specified printer. Each printer driver has a unique string.|
|SpecVersion|An  **Integer** that specifies the version number of the DEVMODE structure in the Win32 Software Development Kit.|
|DriverVersion|An  **Integer** that specifies the printer driver version number assigned by the printer driver developer.|
|Size|An  **Integer** that specifies the size, in bytes, of the DEVMODE structure. (This value doesn't include the optional **dmDriverData** member for device-specific data, which can follow this structure.) If an application manipulates only the driver-independent portion of the data, you can use this member to find out the length of this structure without having to account for different versions.|
|DriverExtra|An  **Integer** that specifies the size, in bytes, of the optional **dmDriverData** member for device-specific data, which can follow this structure. If an application doesn't use device-specific information, you set this member to 0.|
|Fields|A  **Long** value that specifies which of the remaining members in the DEVMODE structure have been initialized.|
|Orientation|An  **Integer** that specifies the orientation of the paper. It can be either 1 (portrait) or 2 (landscape).|
|PaperSize|An  **Integer** that specifies the size of the paper to print on. If you set this member to 0 or 256, the length and width of the paper are specified by the PaperLength and PaperWidth members, respectively. Otherwise, you can set the PaperSize member to a predefined value. For available values, see the[PaperSize member values](values-for-the-papersize-member.md).|
|PaperLength|An  **Integer** that specifies the paper length in units of 1/10 of a millimeter. This member overrides the paper length specified by the PaperSize member for custom paper sizes or for devices such as dot-matrix printers that can print on a variety of paper sizes.|
|PaperWidth|An  **Integer** that specifies the paper width in units of 1/10 of a millimeter. This member overrides the paper width specified by the PaperSize member.|
|Scale|An  **Integer** that specifies the factor by which the printed output will be scaled. The apparent page size is scaled from the physical page size by a factor of _scale_ /100. For example, a piece of paper measuring 8.5 by 11 inches (letter-size) with a Scale value of 50 would contain as much data as a page measuring 17 by 22 inches because the output text and graphics would be half their original height and width.|
|Copies|An  **Integer** that specifies the number of copies printed if the printing device supports multiple-page copies.|
|DefaultSource|An  **Integer** that specifies the default bin from which the paper is fed. For available values, see the[DefaultSource member values](values-for-the-defaultsource-member.md).|
|PrintQuality|An  **Integer** that specifies the printer resolution. The values are ?4 (high), ?3 (medium), ?2 (low), and ?1 (draft).|
|Color|An  **Integer**. For a color printer, specifies whether the output is printed in color. The values are 1 (color) and 2 (monochrome).|
|Duplex|An  **Integer**. For a printer capable of duplex printing, specifies whether the output is printed on both sides of the paper. The values are 1 (simplex), 2 (horizontal), and 3 (vertical).|
|YResolution|An  **Integer** that specifies the y-resolution of the printer in dots per inch (dpi). If the printer initializes this member, the PrintQuality member specifies the x-resolution of the printer in dpi.|
|TTOption|An  **Integer** that specifies how TrueType fonts will be printed. For available values, see the[TTOption member values](values-for-the-ttoption-member.md).|
|Collate|An  **Integer** that specifies whether collation should be used when printing multiple copies. Using uncollated copies provides faster, more efficient output, since the data is sent to the printer just once.|
|FormName|A string with a maximum of 16 characters that specifies the size of paper to use; for example, "Letter" or "Legal".|
|Pad|A  **Long** value that is used to pad out spaces, characters, or values for future versions.|
|Bits|A  **Long** value that specifies in bits per pixel the color resolution of the display device.|
|PW|A  **Long** value that specifies the width, in pixels, of the visible device surface (screen or printer).|
|PH|A  **Long** value that specifies the height, in pixels, of the visible device surface (screen or printer).|
|DFI|A  **Long** value that specifies the device's display mode.|
|DFR|A  **Long** value that specifies the frequency, in hertz (cycles per second), of the display device in a particular mode.|
This property setting is read/write in Design view and read-only in other views.

Printer drivers can add device-specific data immediately following the 94 bytes of the DEVMODE structure. For this reason, it is important that the DEVMODE data outlined above not exceed 94 bytes.

Only printer drivers that export the  **ExtDeviceMode** function use the DEVMODE structure.

An application can retrieve the paper sizes and names supported by a printer by using the DC_PAPERS, DC_PAPERSIZE, and DC_PAPERNAMES values to call the  **DeviceCapabilities** function.

Before setting the value of the TTOption member, applications should find out how a printer driver can use TrueType fonts by using the DC_TRUETYPE value to call the  **DeviceCapabilities** function.


## Example

The following example uses the  **PrtDevMode** property to check the user-defined page size for a report:


```vb
Private Type str_DEVMODE 
 RGB As String * 94 
End Type 
 
Private Type type_DEVMODE 
 strDeviceName As String * 32 
 intSpecVersion As Integer 
 intDriverVersion As Integer 
 intSize As Integer 
 intDriverExtra As Integer 
 lngFields As Long 
 intOrientation As Integer 
 intPaperSize As Integer 
 intPaperLength As Integer 
 intPaperWidth As Integer 
 intScale As Integer 
 intCopies As Integer 
 intDefaultSource As Integer 
 intPrintQuality As Integer 
 intColor As Integer 
 intDuplex As Integer 
 intResolution As Integer 
 intTTOption As Integer 
 intCollate As Integer 
 strFormName As String * 32 
 lngPad As Long 
 lngBits As Long 
 lngPW As Long 
 lngPH As Long 
 lngDFI As Long 
 lngDFr As Long 
End Type 
 
Public Sub CheckCustomPage(ByVal rptName As String) 
 
 Dim DevString As str_DEVMODE 
 Dim DM As type_DEVMODE 
 Dim strDevModeExtra As String 
 Dim rpt As Report 
 Dim intResponse As Integer 
 
 ' Opens report in Design view. 
 DoCmd.OpenReport rptName, acDesign 
 Set rpt = Reports(rptName) 
 
 If Not IsNull(rpt.PrtDevMode) Then 
 strDevModeExtra = rpt.PrtDevMode 
 
 ' Gets current DEVMODE structure. 
 DevString.RGB = strDevModeExtra 
 LSet DM = DevString 
 If DM.intPaperSize = 256 Then 
 
 ' Display user-defined size. 
 intResponse = MsgBox("The current custom page size is " &; _ 
 DM.intPaperWidth / 254 &; " inches wide by " &; _ 
 DM.intPaperLength / 254 &; " inches long. Do you want " &; _ 
 "to change the settings?", vbYesNo + vbQuestion) 
 Else 
 ' Currently not user-defined. 
 intResponse = MsgBox("The report does not have a custom page size. " &; _ 
 "Do you want to define one?", vbYesNo + vbQuestion) 
 End If 
 
 If intResponse = vbYes Then 
 ' User wants to change settings. Initialize fields. 
 DM.lngFields = DM.lngFields Or DM.intPaperSize Or _ 
 DM.intPaperLength Or DM.intPaperWidth 
 
 ' Set custom page. 
 DM.intPaperSize = 256 
 
 ' Prompt for length and width. 
 DM.intPaperLength = InputBox("Please enter page length in inches.") * 254 
 DM.intPaperWidth = InputBox("Please enter page width in inches.") * 254 
 
 ' Update property. 
 LSet DevString = DM 
 Mid(strDevModeExtra, 1, 94) = DevString.RGB 
 rpt.PrtDevMode = strDevModeExtra 
 End If 
 End If 
 
 Set rpt = Nothing 
 
End Sub
```

The following example shows how to change the orientation of the report. This example will switch the orientation from portrait to landscape or landscape to portrait depending on the report's current orientation.




```vb
Public Sub SwitchOrient(ByVal strName As String) 
 
 Const DM_PORTRAIT = 1 
 Const DM_LANDSCAPE = 2 
 Dim DevString As str_DEVMODE 
 Dim DM As type_DEVMODE 
 Dim strDevModeExtra As String 
 Dim rpt As Report 
 
 ' Opens report in Design view. 
 DoCmd.OpenReport strName, acDesign 
 Set rpt = Reports(strName) 
 
 If Not IsNull(rpt.PrtDevMode) Then 
 strDevModeExtra = rpt.PrtDevMode 
 DevString.RGB = strDevModeExtra 
 LSet DM = DevString 
 DM.lngFields = DM.lngFields Or DM.intOrientation 
 
 ' Initialize fields. 
 If DM.intOrientation = DM_PORTRAIT Then 
 DM.intOrientation = DM_LANDSCAPE 
 Else 
 DM.intOrientation = DM_PORTRAIT 
 End If 
 
 ' Update property. 
 LSet DevString = DM 
 Mid(strDevModeExtra, 1, 94) = DevString.RGB 
 rpt.PrtDevMode = strDevModeExtra 
 End If 
 
 Set rpt = Nothing 
 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

