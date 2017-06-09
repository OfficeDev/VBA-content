---
title: Differentiate between Office for Mac versions at compile time
ms.prod: office
ms.date: 06/08/2017
---
# Differentiate between Office for Mac versions at compile time

Use a version conditional to differentiate between Office for Mac 2011 and Office 2016 for Mac.

***Applies to:*** *Excel for Mac | PowerPoint for Mac | Word for Mac | Office 2016 for Mac | Office for Mac 2011*

Office 2016 for Mac supports commands such as [GrantAccessToMultipleFiles](grantaccesstomultiplefiles.md) and [AppleScriptTask](AppleScriptTask.md) that are not supported in other versions of Office. If your solution targets multiple versions of Office, we recommend that you use conditional compilation.  

You can use **MAC_OFFICE_VERSION** to determine which version of VBA the user is running. The following example shows how to use it in your code. 

```vb
    Sub VersionConditionals()

    #If MAC_OFFICE_VERSION >= 15 Then
      Debug.Print "We are running on Mac 15+"
    #Else
      Debug.Print "We are not running on Mac 15+"
    #End If
    #If Mac Then
      Debug.Print "We are running on a Mac"
    #Else
      Debug.Print "We are not running on a Mac"
    #End If
    End Sub
```

**Note:** The "#If Mac" conditional is the same in Office for Mac 2011. 
