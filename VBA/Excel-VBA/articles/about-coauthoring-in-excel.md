---
title: About co-authoring in Excel
ms.prod: EXCEL
ms.assetid: about-coauthoring-in-excel
---


# About co-authoring in Excel

Learn about how co-authoring works in Excel and how to you may need to adjust your add-in/macro for smooth integration with co-authoring.

## About co-authoring

Co-authoring enables you to edit a workbook hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online) simultaneously with other users. With each save, everyone editing the workbook at that time can see changes. If you're not ready for others to see your changes, then you can turn off [AutoSave](../../Office-Shared-VBA/articles/about-autosave.md) until you're ready to share your changes and receive others' changes.

**IMPORTANT**: Previously, add-in/macro code was written with the assumption that the code would be run in the context of a single, local user and didn't have to factor in remote users. However, with co-authoring, if a user runs code which only sets or changes the values of local variable without updating the workbook, then such changes are not pushed to other users.

### **Example scenario**



## Integration with co-authoring events

Co-authoring introduces new events **BeforeRemoteChange** and **AfterRemoteChange** which enable you to handle remote changes.

You may want to receive the latest changes made by another user when the data is fundamental to the expected behavior of the workbook, for example, data visualization and the navigation task pane. 

**Table 1. Examples where the user should receive remote changes**

|**Example**|**Scenario**|
|:-----|:-----|
|Data visualization|Your add-in plots data points on a map based on location data found in a particular range in the workbook. If a user edits any of the location data, all the remote users should receive this change so that each user's map can be updated.|
|Navigation task pane|Your add-in displays all current workbook tabs in a task pane for easy navigation. If a user adds a worksheet, all the remote users should receive this change so that each user's task pane can display a link to the new worksheet.|

### Data visualization example

In this example, we've created a range of locations in London, UK and inserted a Bing map. We are sharing the workbook with someone in a different city. The macro in the first example changes the longitude of Camden Town. With autosave on, the change is passed to the remote user who catches the change with the **AfterRemoteChange** event.

Select cell in location range and change longitude value.
```vb
Sub longitudeChange()

    Range("A5").Select
    ActiveCell.FormulaR1C1 = "51.5390111,-0.1425553"
End Sub

```
```vb
Private Sub Workbook_AfterRemoteChange()
    If Range("A1").Value <> True Then
        Range("A1").Value = True
        'Insert code to modify workbook further
    End If
End Sub
```

Figure 1.
![london locations](images/londonLocations.png) 

```vb
'TODO: do: Visualization
```

### Example

This example enables you to avoid an infinite loop in the **AfterRemoteChange** event.

```vb
Private Sub Workbook_AfterRemoteChange()
    If Range("A1").Value <> True Then
        Range("A1").Value = True
        'Insert code to modify workbook further
    End If
End Sub
```

You may want to avoid changes that cause errors or degraded performance, for example, data validation and data consistency. See Table 2 for example scenarios.

**Table 2. Examples where the user should not receiving remote changes**

|**Example**|**Scenario**|
|:-----|:-----|
|Data validation|A change event is triggered when a specific range is edited in the workbook. Your add-in code then validates the change and, if the check fails, notifies the user via pop-up dialog. However, if all the remote users collaborating on that workbook are notified about a validation failure unrelated to their own changes, this can lead to a poor experience.|
|Data consistency|A change event is triggered and your add-in code synchronizes the data in the workbook with data in another part of the workbook or in an external system. If a remote user receives the change which causes the add-in code to synchronize the same data, this can lead to decreased performance for the remote user or data duplication in the external system.|

### Example

This example enables you to

```vb
'TODO: NOT: Data validation
```

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/about-autosave.md)

#### Additional resources

[Collaborate on Excel workbooks at the same time with co-authoring](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)
