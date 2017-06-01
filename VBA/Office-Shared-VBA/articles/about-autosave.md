---
title: About AutoSave in Office
ms.prod: MULTIPLEPRODUCTS
ms.assetid: about-autosave
---


# About AutoSave in Office

Learn about how AutoSave works in Excel, PowerPoint, and Word and how it can impact add-in/macro integration. For more about how AutoSave works in general, see [this link](https://support.office.com/en-US/article/What-is-AutoSave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5).

## About AutoSave

When a file is hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), AutoSave enables the user's edits to be saved automatically and continuously. When the file is shared with others, then the others' changes will be merged into this user's version of the file. If AutoSave is turned off, then save must be triggered manually for the user's changes to be persisted in the cloud and for that user to receive others' changes.

## Occasions where AutoSave is disabled

If a file is not hosted in the cloud but is instead saved elsewhere (for example, on your local machine), then AutoSave is disabled.

## Known issues and mitigations

In certain cases, the integration between AutoSave and your add-in/macro's subscription to **BeforeSave**/**AfterSave** events may lead to a degraded experience for your users. In the following table, you will find a number of those cases and possible mitigations. For each issue, you can turn off AutoSave or adjust your code according to the mitigation.

**Table 1. Known issues and possible mitigations**

|**Issue**|**Example scenario**|**Possible mitigation**|
|:-----|:-----|:-----|
|Save events take too long||Make the events more efficient|
|Save events display a modal dialog||Remove modal dialog|
|(Excel only) Save events clear the undo stack||Turn off AutoSave and notify users by email or other communication to save manually|
|(Excel only) **AfterSave** event edits the file which leads to repeated attempts to save||Remove the edit if it is not needed|
|**BeforeSave** event cancels the file save which leads to repeated attempts to save||Do not cancel the file save if it is unnecessary|

## Example

This example turns off AutoSave and notifies the user that the workbook is not being saved automatically.

```vb
Sub UseAutoSaveOn()
    ActiveWorkbook.autoSaveOn = False
    MsgBox "This workbook is being saved automatically: " & ActiveWorkbook.autoSaveOn
End Sub
```

## See also

#### Concepts

[Co-authoring](about-coauthoring-in-excel.md)

[Workbook Object](workbook-object-excel.md)

#### Additional resources

[What is AutoSave?](https://support.office.com/en-US/article/What-is-AutoSave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5)