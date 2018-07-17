---
title: Run an AppleScript with VB
ms.prod: office
ms.date: 06/08/2017
---
# Run an AppleScript with VB
 
Call an AppleScript file from a VB macro in Office 2016 for Mac.

***Applies to:*** *Excel for Mac | PowerPoint for Mac | Word for Mac | Office 2016 for Mac*

The **AppleScriptTask** command executes an AppleScript script file located outside the sandboxed app. 

The following code shows how to call **AppleScriptTask** from VB.

```vb
    Dim myScriptResult as String
    myScriptResult = AppleScriptTask ("MyAppleScriptFile.applescript", "myapplescripthandler", "my parameter string") 
```

The MyAppleScript.applescript file must be in ~/Library/Application Scripts/[bundle id]/. The .applescript extension is not required; you can also use the .scpt extension.

“Myapplescripthandler” is the name of a script handler in the MyAppleScript.applescript file.
“My parameter string” is the single input parameter to the “myapplescripthandler” script handler.

The following are the [bundle id] values for Excel, PowerPoint, and Word:

- com.microsoft.Word
- com.microsoft.Excel
- com.microsoft.Powerpoint

For example, the corresponding AppleScript for Excel would be in a file named "MyAppleScriptFile.applescript" that is in ~/Library/Application Scripts/com.microsoft.Excel/.

Remember: The folders com.microsoft.Excel etc. may not exist. In that case, just create them using standard "mk dir" command. 
The following is an example of a handler.

```vb
    on myapplescripthandler(paramString) 

    #do something with paramString 
    return "You told me " & paramString 

    end myapplescripthandler
```

##What happened to MacScript?
Earlier versions of Office for Mac implemented a command called **MacScript** that supported inline AppleScripts. Although that command still exists in Office 2016 for Mac, **MacScript** is deprecated. Due to sandbox restrictions, the **MacScript** command cannot invoke other applications, such as Finder, in Office 2016 for Mac. We recommend that you use the **AppleScriptTask** command instead of the **MacScript** command in apps for Office 2016 for Mac.
