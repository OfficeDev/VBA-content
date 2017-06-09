---
title: AllowBypassKey Property
keywords: vbaac10.chm10101
f1_keywords:
- vbaac10.chm10101
ms.prod: access
api_name:
- Access.AllowBypassKey
ms.assetid: fa693699-f96b-b287-5248-828e9be1bbbe
ms.date: 06/08/2017
---


# AllowBypassKey Property

  

**Applies to:** Access 2013 | Access 2016

 **In this article**
[Setting](#sectionSection0)
[Remarks](#sectionSection1)
[ Example](#sectionSection2)


You can use the  **AllowBypassKey** property to specify whether or not the SHIFT key is enabled for bypassing the startup properties and the AutoExec macro. For example, you can set the **AllowBypassKey** property to **False** to prevent a user from bypassing the startup properties and the AutoExec macro.

## Setting
<a name="sectionSection0"> </a>

The  **AllowBypassKey** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|Enable the SHIFT key to allow the user to bypass the startup properties and the AutoExec macro.|
|**False**|Disable the SHIFT key to prevent the user from bypassing the startup properties and the AutoExec macro.|
You can set this property by using a macro or Visual Basic .

To set the  **AllowBypassKey** property by using a macro or Visual Basic, you must create the property in the following ways:


- In a Microsoft Access database, you can add it by using the  **[CreateProperty](http://msdn.microsoft.com/library/F2039BE9-5FD8-F673-DFBF-0A71540CDC98%28Office.15%29.aspx)** method and append it to the **Properties** collection of the **Database** object.
    
- In a Microsoft Access project (.adp), you can add it to the  **[AccessObjectProperties](http://msdn.microsoft.com/library/2df86891-6038-d147-2a32-f1c77b841067%28Office.15%29.aspx)** collection of the **[CurrentProject](http://msdn.microsoft.com/library/e6baae73-1eeb-b48f-d35e-b3e921378561%28Office.15%29.aspx)** object by using the **[Add](http://msdn.microsoft.com/library/8f86d5f8-b9af-87d3-fae4-e1a24d7225b6%28Office.15%29.aspx)** method.
    

## Remarks
<a name="sectionSection1"> </a>

You should make sure the  **AllowBypassKey** property is set to **True** when debugging an application.

This property's setting doesn't take effect until the next time the application database opens.


## Example
<a name="sectionSection2"> </a>

The following example shows a procedure named SetBypassProperty that passes the name of the property to be set, its data type, and its desired setting. The general purpose procedure ChangeProperty attempts to set the  **AllowBypassKey** property and, if the property isn't found, uses the **CreateProperty** method to append it to the **Properties** collection. This is necessary because this property doesn't appear in the **Properties** collection until its been added.


```vb
Sub SetBypassProperty() 
Const DB_Boolean As Long = 1 
    ChangeProperty "AllowBypassKey", DB_Boolean, False 
End Sub 
 
Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer 
    Dim dbs As Object, prp As Variant 
    Const conPropNotFoundError = 3270 
 
    Set dbs = CurrentDb 
    On Error GoTo Change_Err 
    dbs.Properties(strPropName) = varPropValue 
    ChangeProperty = True 
 
Change_Bye: 
    Exit Function 
 
Change_Err: 
    If Err = conPropNotFoundError Then    ' Property not found. 
        Set prp = dbs.CreateProperty(strPropName, _ 
            varPropType, varPropValue) 
        dbs.Properties.Append prp 
        Resume Next 
    Else 
        ' Unknown error. 
        ChangeProperty = False 
        Resume Change_Bye 
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

