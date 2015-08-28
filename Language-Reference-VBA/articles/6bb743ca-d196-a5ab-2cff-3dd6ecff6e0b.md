
# FreeSpace Property

 **Last modified:** July 28, 2015


 **Description**
Returns the amount of free space available to a user on the specified drive or network share. Read-only.
 **Syntax**
 _object_. **FreeSpace**
The  _object_ is always a **Drive** object.
 **Remarks**
The value returned by the  **FreeSpace** property is typically the same as that returned by the **AvailableSpace** property. Differences may occur between the two for computer systems that support quotas.
The following code illustrates the use of the  **FreeSpace** property:



```
Sub ShowFreeSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " &amp; UCase(drvPath) &amp; " - " 
    s = s &amp; d.VolumeName  &amp; vbCrLf
    s = s &amp; "Free Space: " &amp; FormatNumber(d.FreeSpace/1024, 0) 
    s = s &amp; " Kbytes"
    MsgBox s
End Sub

```

