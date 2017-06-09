---
title: VolumeName Property
keywords: vblr6.chm2181965
f1_keywords:
- vblr6.chm2181965
ms.prod: office
api_name:
- Office.VolumeName
ms.assetid: 8592ae63-f36e-e87a-8286-72419d7781d0
ms.date: 06/08/2017
---


# VolumeName Property



 **Description**
Sets or returns the volume name of the specified drive. Read/write.
 **Syntax**
 _object_. **VolumeName** [= _newname_ ]
The VolumeName property has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **Drive** object.|
| _newname_|Optional. If provided,  _newname_ is the new name of the specified _object_.|
 **Remarks**
The following code illustrates the use of the  **VolumeName** property:



```vb
Sub ShowVolumeInfo(drvpath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = "Drive " &; d.DriveLetter &; ": - " &; d.VolumeName
    MsgBox s
End Sub
```


