---
title: CurrentProject.IsTrusted Property (Access)
keywords: vbaac10.chm12730
f1_keywords:
- vbaac10.chm12730
ms.prod: access
api_name:
- Access.CurrentProject.IsTrusted
ms.assetid: c3d8b6f8-c79f-79ab-d4e0-0454f97ac937
ms.date: 06/08/2017
---


# CurrentProject.IsTrusted Property (Access)

Gets whether or not macros and Visual Basic for Applications (VBA) code have been enabled in the current project. Read-only  **Boolean**.


## Syntax

 _expression_. **IsTrusted**

 _expression_ A variable that represents a **CurrentProject** object.


## Example

The following example shows how to use the  **IsTrusted** property in a macro to determine whether the database has been opened with trust enabled. If trust has been enabled, the Visual Basic for Applications (VBA) subroutine **Init** is called. Otherwise, the use is notified that the database has been opened in disabled mode.

 **Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)


```text
If [currentproject].[istrusted] Then
    RunCode
        Function Name =Init()

Else
    MessageBox
        Message The application is opened in disabled mode. Please enable the application for full functionality.
        Beep Yes
        Type None
        Title Disabled Mode Check

End If
```


## About the Contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[CurrentProject Object](currentproject-object-access.md)

