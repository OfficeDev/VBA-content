---
title: Workbook.BeforeRemoteChange Event (Excel)
keywords: vbaxl10.chm504120
f1_keywords:
- vbaxl10.chm504120
ms.prod: excel
api_name:
- Excel.Workbook.BeforeRemoteChange
ms.date: 06/08/2017
---


# Workbook.BeforeRemoteChange Event (Excel)

Occurs before a remote user's edits to the workbook are merged.

## Syntax

_expression_.**BeforeRemoteChange**

_expression_ A variable that represents a **Workbook** object.


## Parameters

None

## Return value

Nothing

## Example

This example notifies the user that there is an incoming remote change.

```vb
Private Sub Workbook_BeforeRemoteChange()
    'A remote user has made a change to this workbook.
    'The code in this subroutine will be run before those changes are merged.
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[co authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
