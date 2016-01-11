# Run an AppleScript with VB
 
Call an AppleScript file from a VB macro in Office 2016 for Mac.

**Last modified:** January 11, 2016 

***Applies to:*** *Excel for Mac | PowerPoint for Mac | Word for Mac | Office 2016 for Mac*

The **AppleScriptTask** command executes an AppleScript script file located outside the sandboxed app. 

The following code shows how to call **AppleScriptTask** from VB.

```
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

The following is an example of a handler.

```
on myapplescripthandler(paramString) 

    #do something with paramString 
    return "You told me " & paramString 

end myapplescripthandler
```

##What happened to MacScript?
The **MacScript** command that supports inline AppleScripts in Office for Mac 2011 is deprecated. Due to sandbox restrictions, the **MacScript** command cannot invoke other applications, such as Finder, in Office 2016 for Mac. We recommend that you use the **AppleScriptTask** command instead of the **MacScript** command in apps for Office 2016 for Mac. 
