
# InvisibleApp.HelpPaths Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Gets or sets the paths where Microsoft Visio looks for Help files. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **HelpPaths**

 _expression_A variable that represents an  **InvisibleApp** object.


### Return Value

String


## Remarks
<a name="sectionSection1"> </a>

The  **HelpPaths** property is set to an empty string ("") by default.

The string passed to and received from the  **HelpPaths** property is the same string shown in the **File Paths** dialog box. (Click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click ** File Locations**.) This string is stored in the  **HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Visio\Application\HelpPath** subkey.

When the application looks for Help files, it looks in all paths named in the  **HelpPaths** property and all the subfolders of those paths. If you pass the **HelpPaths** property to the **EnumDirectories** method, it returns a complete list of fully qualified paths in the folders passed in.

Setting the  **HelpPaths** property replaces existing values for **HelpPaths** in the **File Paths** dialog box. To retain existing values, get the existing string and then append the new file path to that string, as shown in the following code:




```
Application.HelpPaths = Application.HelpPaths &amp; ";" &amp; "newpath".
```


 **Note**  Modifying the registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


## Example
<a name="sectionSection2"> </a>

This Microsoft Visual Basic for Applications (VBA) macro shows how to get and set the  **HelpPaths** property of the **Application** object. Before running this macro, replace _fullpath(s)_ with the path or paths to the location or locations where you want Visio to look for Help files.


```
 
Public Sub GetHelpPaths_Example()  
 
    Dim strCurrentPath As String 
 
    'Retrieve the current path to Visio Help files.  
    strCurrentPath = Application.HelpPaths  
    MsgBox ("The current path for Microsoft Visio Help files is " + strCurrentPath)  
 
End Sub   
 
Public Sub SetHelpPaths_Example()  
 
    Dim strNewPath As String 
 
    'Store the new path.  
    strNewPath = "fullpath(s)"  
 
    'Set the new path in the Application object.  
    Application.HelpPaths = strNewPath  
 
End Sub 

```

