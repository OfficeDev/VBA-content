---
title: Application.DBEngine Property (Access)
keywords: vbaac10.chm12545
f1_keywords:
- vbaac10.chm12545
ms.prod: access
api_name:
- Access.Application.DBEngine
ms.assetid: ad4638e4-0c72-ce24-e322-e147e2f0cfc2
ms.date: 06/08/2017
---


# Application.DBEngine Property (Access)

You can use the  **DBEngine** property in[Visual Basic](set-properties-by-using-visual-basic.md)to access the current  **DBEngine** object and its related properties. Read-only **DBEngine**.


## Syntax

 _expression_. **DBEngine**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **DBEngine** property of the **[Application](application-object-access.md)** object represents the Microsoft Access database engine. The **DBEngine** object is the top-level object in the Data Access Objects (DAO) model and it contains and controls all other objects in the hierarchy of Data Access Objects.


## Example

The following example displays the  **DBEngine** properties in a message box.


```vb
Private Sub Command1_Click() 
 DisplayApplicationInfo Me 
End Sub 
 
Function DisplayApplicationInfo(obj As Object) As Integer 
 Dim objApp As Object, intI As Integer, strProps As String 
 On Error Resume Next 
 ' Form Application property. 
 Set objApp = obj.Application 
 MsgBox "Application Visible property = " &; objApp.Visible 
 If objApp.UserControl = True Then 
 For intI = 0 To objApp.DBEngine.Properties.Count - 1 
 strProps = strProps &; objApp.DBEngine.Properties(intI).Name &; ", " 
 Next intI 
 End If 
 MsgBox Left(strProps, Len(strProps) - 2) &; ".", vbOK, "DBEngine Properties" 
End Function
```


## See also


#### Concepts


[Application Object](application-object-access.md)

