---
title: Programmatically Retrieve Printer Capabilities
ms.prod: access
ms.assetid: 8c929823-6b61-16ea-6d84-ff47cc1e8389
ms.date: 06/08/2017
---


# Programmatically Retrieve Printer Capabilities

The  **[Printers](printers-object-access.md)** collection and **[Printer](printer-object-access.md)** object allow you only to set or retrieve settings for a printer. To determine a printer's capabilities, such as the kinds of paper or paper bins it supports, you must use calls to the Windows Application Programming Interface (API) **DeviceCapabilities** function. It is beyond the scope of this topic to cover this in detail, but the following code sample from the modPrinters module of the PrinterDemo.mdb sample download demonstrates how to retrieve the names and IDs of the supported paper size and paper bins for a printer.

The following code should be pasted into the general declarations section of a module.



```vb
' Declaration for the DeviceCapabilities function API call. 
Private Declare Function DeviceCapabilities Lib "winspool.drv" _ 
    Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, _ 
    ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _ 
    ByVal lpDevMode As Long) As Long 
     
' DeviceCapabilities function constants. 
Private Const DC_PAPERNAMES = 16 
Private Const DC_PAPERS = 2 
Private Const DC_BINNAMES = 12 
Private Const DC_BINS = 6 
Private Const DEFAULT_VALUES = 0 

```

The following procedure uses the  **DeviceCapabilities** API function to display a message box with the name of the default printer and a list of the paper sizes it supports.



```vb
Sub GetPaperList() 
    Dim lngPaperCount As Long 
    Dim lngCounter As Long 
    Dim hPrinter As Long 
    Dim strDeviceName As String 
    Dim strDevicePort As String 
    Dim strPaperNamesList As String 
    Dim strPaperName As String 
    Dim intLength As Integer 
    Dim strMsg As String 
    Dim aintNumPaper() As Integer 
     
    On Error GoTo GetPaperList_Err 
     
    ' Get the name and port of the default printer. 
    strDeviceName = Application.Printer.DeviceName 
    strDevicePort = Application.Printer.Port 
     
    ' Get the count of paper names supported by the printer. 
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_PAPERNAMES, _ 
        lpOutput:=ByVal vbNullString, _ 
        lpDevMode:=DEFAULT_VALUES) 
     
    ' Re-dimension the array to the count of paper names. 
    ReDim aintNumPaper(1 To lngPaperCount) 
     
    ' Pad the variable to accept 64 bytes for each paper name. 
    strPaperNamesList = String(64 * lngPaperCount, 0) 
 
    ' Get the string buffer of all paper names supported by the printer. 
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_PAPERNAMES, _ 
        lpOutput:=ByVal strPaperNamesList, _ 
        lpDevMode:=DEFAULT_VALUES) 
     
    ' Get the array of all paper numbers supported by the printer. 
    lngPaperCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_PAPERS, _ 
        lpOutput:=aintNumPaper(1), _ 
        lpDevMode:=DEFAULT_VALUES) 
     
    ' List the available paper names. 
    strMsg = "Papers available for " &; strDeviceName &; vbCrLf 
    For lngCounter = 1 To lngPaperCount 
         
        ' Parse a paper name from the string buffer. 
        strPaperName = Mid(String:=strPaperNamesList, _ 
            Start:=64 * (lngCounter - 1) + 1, Length:=64) 
        intLength = VBA.InStr(Start:=1, String1:=strPaperName, String2:=Chr(0)) - 1 
        strPaperName = Left(String:=strPaperName, Length:=intLength) 
         
        ' Add a paper number and name to text string for the message box. 
        strMsg = strMsg &; vbCrLf &; aintNumPaper(lngCounter) _ 
            &; vbTab &; strPaperName 
             
    Next lngCounter 
         
    ' Show the paper names in a message box. 
    MsgBox Prompt:=strMsg 
 
GetPaperList_End: 
    Exit Sub 
     
GetPaperList_Err: 
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical &; vbOKOnly, _ 
        Title:="Error Number " &; Err.Number &; " Occurred" 
    Resume GetPaperList_End 
     
End Sub
```

The following procedure uses the  **DeviceCapabilities** API function to display a message box with the name of the default printer and a list of the paper bins it supports.



```vb
Sub GetBinList(strName As String) 
' Uses the DeviceCapabilities API function to display a 
' message box with the name of the default printer and a 
' list of the paper bins it supports. 
 
    Dim lngBinCount As Long 
    Dim lngCounter As Long 
    Dim hPrinter As Long 
    Dim strDeviceName As String 
    Dim strDevicePort As String 
    Dim strBinNamesList As String 
    Dim strBinName As String 
    Dim intLength As Integer 
    Dim strMsg As String 
    Dim aintNumBin() As Integer 
     
    On Error GoTo GetBinList_Err 
     
    ' Get name and port of the default printer. 
    strDeviceName = Application.Printers(strName).DeviceName 
    strDevicePort = Application.Printers(strName).Port 
     
    ' Get count of paper bin names supported by the printer. 
    lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_BINNAMES, _ 
        lpOutput:=ByVal vbNullString, _ 
        lpDevMode:=DEFAULT_VALUES) 
     
    ' Re-dimension the array to count of paper bins. 
    ReDim aintNumBin(1 To lngBinCount) 
     
    ' Pad variable to accept 24 bytes for each bin name. 
    strBinNamesList = String(Number:=24 * lngBinCount, Character:=0) 
 
    ' Get string buffer of paper bin names supported by the printer. 
    lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_BINNAMES, _ 
        lpOutput:=ByVal strBinNamesList, _ 
        lpDevMode:=DEFAULT_VALUES) 
         
    ' Get array of paper bin numbers supported by the printer. 
    lngBinCount = DeviceCapabilities(lpsDeviceName:=strDeviceName, _ 
        lpPort:=strDevicePort, _ 
        iIndex:=DC_BINS, _ 
        lpOutput:=aintNumBin(1), _ 
        lpDevMode:=0) 
         
    ' List available paper bin names. 
    strMsg = "Paper bins available for " &; strDeviceName &; vbCrLf 
    For lngCounter = 1 To lngBinCount 
         
        ' Parse a paper bin name from string buffer. 
        strBinName = Mid(String:=strBinNamesList, _ 
            Start:=24 * (lngCounter - 1) + 1, _ 
            Length:=24) 
        intLength = VBA.InStr(Start:=1, _ 
            String1:=strBinName, String2:=Chr(0)) - 1 
        strBinName = Left(String:=strBinName, _ 
                Length:=intLength) 
 
        ' Add bin name and number to text string for message box. 
        strMsg = strMsg &; vbCrLf &; aintNumBin(lngCounter) _ 
            &; vbTab &; strBinName 
             
    Next lngCounter 
         
    ' Show paper bin numbers and names in message box. 
    MsgBox Prompt:=strMsg 
     
GetBinList_End: 
    Exit Sub 
GetBinList_Err: 
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical &; vbOKOnly, _ 
        Title:="Error Number " &; Err.Number &; " Occurred" 
    Resume GetBinList_End 
End Sub
```


