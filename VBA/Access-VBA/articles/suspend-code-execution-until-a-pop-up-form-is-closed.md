---
title: Suspend Code Execution Until a Pop-up Form is Closed
ms.prod: access
ms.assetid: d4d419ac-bf43-3356-4c20-e9bb74f9f591
ms.date: 06/08/2017
---


# Suspend Code Execution Until a Pop-up Form is Closed

To ensure that code in a form suspends operation until a pop-up form is closed, you must open the pop-up form as a modalwindow. The following example illustrates how to use the  **[OpenForm](docmd-openform-method-access.md)** method to do this.


```
doCmd.OpenForm FormName:=<Name of form to open>, WindowMode:=acDialog
```


