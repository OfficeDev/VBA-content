---
title: AppIcon Property
keywords: vbaac10.chm10110
f1_keywords:
- vbaac10.chm10110
ms.prod: access
api_name:
- Access.AppIcon
ms.assetid: e322784a-39f4-0055-c15e-5051a382c68e
ms.date: 06/08/2017
---


# AppIcon Property

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[Example](#sectionSection2)


You can use the  **AppIcon** property to specify the name of the bitmap (.bmp) or icon (.ico) file that contains the application's icon. For example, you can use the **AppIcon** property to specify a .bmp file that contains a picture of an automobile to represent an automotive parts application.

## Setting
<a name="sectionSection0"> </a>

The  **AppIcon** property is a string expression that's a valid bitmap or icon file name (/including the path).

The easiest way to set this property is by using the  **Application Icon** option in the **Access Options** dialog box, available by clicking the click the **Microsoft Office Button**
![File menu button](images/O12FileMenuButton_ZA10077102.gif) and then clicking the **Current Database** category. You can also set this property by using a macro or Visual Basic .

To set the  **AppIcon** property by using a macro or Visual Basic, you must first either set the property in the **Access Options** dialog box once or create the property in the following ways:


- In a Microsoft Access database , you can add it by using the  **CreateProperty** method and append it to the **Properties** collection of the **Database** object.
    
- In a Microsoft Access project (.adp), you can add it to the  **AccessObjectProperties** collection of the **CurrentProject** object by using the **Add** method.
    
You must also use the  **RefreshTitleBar** method to make any changes visible immediately.


## Remarks
<a name="sectionSection1"> </a>

If you are distributing your application, it's recommended that the .bmp or .ico file containing the icon reside in the same directory as your Microsoft Access application.

If the  **AppIcon** property isn't set or is invalid, the Microsoft Access icon is displayed.

This property setting takes effect immediately after it's set in code (as long as the code includes the  **RefreshTitleBar** method) or the **Access Options** dialog box is closed.


## Example
<a name="sectionSection2"> </a>

The following example shows how to change the  **AppIcon** and **AppTitle** properties in a Microsoft Access database. If the properties haven't already been set or created, you must create them and append them to the **Properties** collection by using the **CreateProperty** method.


```vb
Sub cmdAddProp_Click() 
 Dim intX As Integer 
 Const DB_Text As Long = 10 
 intX = AddAppProperty("AppTitle", DB_Text, "My Custom Application") 
 intX = AddAppProperty("AppIcon", DB_Text, "C:\Windows\Cars.bmp") 
 CurrentDb.Properties("UseAppIconForFrmRpt") = 1 
 Application.RefreshTitleBar 
End Sub 
 
Function AddAppProperty(strName As String, _ 
 varType As Variant, varValue As Variant) As Integer 
 Dim dbs As Object, prp As Variant 
 Const conPropNotFoundError = 3270 
 
 Set dbs = CurrentDb 
 On Error GoTo AddProp_Err 
 dbs.Properties(strName) = varValue 
 AddAppProperty = True 
 
AddProp_Bye: 
 Exit Function 
 
AddProp_Err: 
 If Err = conPropNotFoundError Then 
 Set prp = dbs.CreateProperty(strName, varType, varValue) 
 dbs.Properties.Append prp 
 Resume 
 Else 
 AddAppProperty = False 
 Resume AddProp_Bye 
 End If 
End Function
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

