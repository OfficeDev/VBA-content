---
title: AccessObject.IsLoaded Property (Access)
keywords: vbaac10.chm12750
f1_keywords:
- vbaac10.chm12750
ms.prod: access
api_name:
- Access.AccessObject.IsLoaded
ms.assetid: 5e68398c-8a95-f3e1-87ec-e2d637f34429
ms.date: 11/30/2017
---


# AccessObject.IsLoaded Property (Access)

You can use the **IsLoaded** property to determine if an **[AccessObject](accessobject-object-access.md)** is currently loaded. Read-only **Boolean**.


## Syntax

 _expression_. **IsLoaded**

 _expression_ A variable that represents an **AccessObject** object.


## Remarks

The **IsLoaded** property uses the following settings.

|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The specified **AccessObject** is loaded.|
|No|**False**|The specified **AccessObject** is not loaded.|

## Example

The following example shows how to prevent a user from opening a particular form directly from the navigation pane.


**Sample code provided by:** The [Microsoft Access 2010 Programmer's Reference](http://www.wrox.com/WileyCDA/WroxTitle/Access-2010-Programmer-s-Reference.productCd-0470591668.html)

```vb
'Don't let this form be opened from the Navigator
If Not CurrentProject.AllForms(cFormUsage).IsLoaded Then
    MsgBox "This form cannot be opened from the Navigation Pane.", _
        vbInformation + vbOKOnly, "Invalid form usage"
    Cancel = True
    Exit Sub
End If
```

## About the contributors
<a name="AboutContributors"> </a>

Wrox Press is driven by the Programmer to Programmer philosophy. Wrox books are written by programmers for programmers, and the Wrox brand means authoritative solutions to real-world programming problems. 

## See also

[AccessObject Object](accessobject-object-access.md)

