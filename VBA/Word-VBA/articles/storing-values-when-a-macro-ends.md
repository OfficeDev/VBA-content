---
title: Storing Values When a Macro Ends
ms.prod: word
ms.assetid: 25d6d0ba-d103-3573-d534-1157cfb603f0
ms.date: 06/08/2017
---


# Storing Values When a Macro Ends

When a macro ends, the values stored in its variables are not automatically saved to disk. If a macro needs to preserve a value, it must store that value outside itself before the macro execution is completed. This topic describes five locations where macro values can be easily stored and retrieved.


## Document variables

Document variables allow you to store values as part of a document or a template. For example, you might store macro values in the document or template where the macro resides. You can add variables to a document or template using the  **[Add](variables-add-method-word.md)** method of the  **[Variables](variables-object-word.md)** collection. The following example saves a document variable in the same location as the macro that is running (document or template) using the  **ActiveDocument** property.


```vb
Sub AddDocumentVariable() 
 ActiveDocument.Variables.Add Name:="Age", Value:=12 
End Sub
```

The following example uses the  **[Value](variable-value-property-word.md)** property with a  **[Variable](variable-object-word.md)** object to return the value of a document variable.




```vb
Sub UseDocumentVariable() 
 Dim intAge As Integer 
 intAge = ActiveDocument.Variables("Age").Value 
End Sub
```


## Remarks

You can use the DOCVARIABLE field to insert a document variable into a document.


## Document properties

Like document variables, document properties allow you to store values as part of a document or a template. Document properties can be viewed in the  **Properties** dialog box.

The Word object model breaks document properties into two groups: built-in and custom. Custom document properties include the properties shown on the  **Custom** tab in the **Properties** dialog box. Built-in document properties include the properties on all the tabs in the **Properties** dialog box, except the **Custom** tab.

To access built-in properties, use the  **BuiltInDocumentProperties** property to return a **DocumentProperties** collection that includes the built-in document properties. Use the **CustomDocumentProperties**property of a  **[Document](document-object-word.md)** object or a **[Template](template-object-word.md)** object to return a **DocumentProperties** collection that includes the custom document properties. The following example creates a custom document property named "YourName" in the same location as the macro that is running (document or template).




```vb
Sub AddCustomDocumentProperties() 
 ActiveDocument.CustomDocumentProperties.Add Name:="YourName", _ 
 LinkToContent:=False, Value:="Joe", Type:=msoPropertyTypeString 
End Sub
```

Built-in document properties cannot be added to the  **DocumentProperties** collection returned by the **BuiltInDocumentProperties** property of a **[Document](document-object-word.md)** object or **[Template](template-object-word.md)** object. You can, however, retrieve the contents of a built-in document property or change the value of a read/write built-in document property.


## Remarks

You can use the DOCPROPERTY field to insert document properties into a document.


## AutoText entries

AutoText entries can be used to store information in a template. Unlike a document variable or property, AutoText entries can include items beyond macro variables, such as formatted text or a graphic. Use the  **[Add](autotextentries-add-method-word.md)** method with the  **[AutoTextEntries](autotextentries-object-word.md)** collection to create a new AutoText entry. The following example creates an AutoText entry named "MyText" that contains the contents of the selection. If the following instruction is part of a template macro, the new AutoText entry is stored in the template, otherwise, the AutoText entry is stored in the template attached to the document where the instruction resides.


```vb
Sub AddAutoTextEntry() 
 ActiveDocument.AttachedTemplate.AutoTextEntries.Add Name:="MyText", _ 
 Range:=Selection.Range 
End Sub
```

Use the  **[Value](autotextentry-value-property-word.md)** property with an  **[AutoTextEntry](autotextentry-object-word.md)** object to retrieve the contents of an AutoText entry object.


## Settings files

You can set and retrieve information from a settings file using the  **[PrivateProfileString](system-privateprofilestring-property-word.md)** property of the  **[System](system-object-word.md)** object. The structure of a Windows settings file is the same as the Windows 3.1 WIN.INI file. The following example sets the DocNum key to 1 under the DocTracker section in the Macro.ini file.


```vb
Sub MacroSystemFile() 
 System.PrivateProfileString( _ 
 FileName:="C:\My Documents\Macro.ini", _ 
 Section:="DocTracker", Key:="DocNum") = 1 
End Sub
```

After running the above instruction, the Macro.ini file includes the following text.




```
[DocTracker] 
DocNum=1
```

The  **PrivateProfileString** property has three arguments: **_FileName_**,  **_Section_**, and  **_Key_**. The  **_FileName_** argument is used to specify a settings file path and file name. The **_Section_** argument specifies the section name that appears between brackets before the associated keys (do not include the brackets with the section name). The **_Key_** argument specifies the key name, which is followed by an equal sign (=) and the setting.

Use the same  **PrivateProfileString** property to retrieve a setting from a settings file. The following example retrieves the DocNum setting under the DocTracker section in the Macro.ini file.




```vb
Sub GetSystemFileInfo() 
 Dim intDocNum As Integer 
 intDocNum = System.PrivateProfileString( _ 
 FileName:="C:\My Documents\Macro.ini", _ 
 Section:="DocTracker", Key:="DocNum") 
 MsgBox "DocNum is " &; intDocNum 
End Sub
```


## Windows registry

You can set and retrieve information from the Windows registry using the  **PrivateProfileString**property. The following example retrieves the Word 2007 program directory from the Windows registry.


```vb
Sub GetRegistryInfo() 
 Dim strSection As String 
 Dim strPgmDir As String 
 strSection = "HKEY_CURRENT_USER\Software\Microsoft" _ 
 &; "\Office\12.0\Word\Options" 
 strPgmDir = System.PrivateProfileString(FileName:="", _ 
 Section:=strSection, Key:="PROGRAMDIR") 
 MsgBox "The directory for Word is - " &; strPgmDir 
End Sub
```

The  **PrivateProfileString** property has three arguments: **_FileName_**,  **_Section_**, and  **_Key_**. To return or set a value for a registry entry, specify an empty string ("") for the  **_FileName_** argument. The **_Section_** argument should be the complete path to the registry subkey. The **_Key_** argument should be the name of an entry in the subkey specified by **_Section_**.

You can also set information in the Windows registry using the following  **PrivateProfileString** syntax.

 **System.PrivateProfileString**( **_FileName_**,  **_Section_**,  **_Key_**)  **=**_value_

The following example sets the DOC-PATH entry to "C:\My Documents" in the Options subkey for Office Word 2007 in the Windows registry.




```vb
Sub SetDocumentDirectory() 
 Dim strDocDirectory As String 
 strDocDirectory = "HKEY_CURRENT_USER\Software\Microsoft" _ 
 &; "\Office\10.0\Word\Options" 
 System.PrivateProfileString(FileName:="", _ 
 Section:=strDocDirectory, Key:="DOC-PATH") = "C:\My Documents" 
End Sub
```


