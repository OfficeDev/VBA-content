
# FreeSpace Property



 **Description**
Returns the amount of free space available to a user on the specified drive or network share. Read-only.
 **Syntax**
 _object_. **FreeSpace**
The  _object_ is always a **Drive** object.
 **Remarks**
The value returned by the  **FreeSpace** property is typically the same as that returned by the **AvailableSpace** property. Differences may occur between the two for computer systems that support quotas.
The following code illustrates the use of the  **FreeSpace** property:



```vb
Sub ShowFreeSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " &; UCase(drvPath) &; " - " 
    s = s &; d.VolumeName  &; vbCrLf
    s = s &; "Free Space: " &; FormatNumber(d.FreeSpace/1024, 0) 
    s = s &; " Kbytes"
    MsgBox s
End Sub
```

