---
title: Workbook.AfterRemoteChange Event (Excel)
keywords: vbaxl10.chm504121
f1_keywords:
- vbaxl10.chm504121
ms.prod: excel
api_name:
- Excel.Workbook.AfterRemoteChange
ms.date: 06/08/2017
---


# Workbook.AfterRemoteChange Event (Excel)

Occurs after a remote user's edits to the workbook are merged.

## Syntax

_expression_.**AfterRemoteChange**

_expression_ A variable that represents a Workbook object.

## Parameters

None

## Return value

Nothing

## Example

This example notifies the user that there was an incoming remote change.

```vb
Private Sub Workbook_AfterRemoteChange()
    'A remote user has made a change to this workbook and that change has been merged.
    'The code in this subroutine will now be run.
End Sub
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

[coauthoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)
