---
title: About co-authoring in Excel
ms.prod: excel
ms.assetid: about-coauthoring-in-excel
ms.date: 06/08/2017
---


# About co-authoring in Excel

Learn about how co-authoring works in Excel 2016 for O365 subscribers and how you may need to adjust your add-in/macro for smooth integration with co-authoring.

## About co-authoring

Co-authoring enables you to edit a workbook hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online) simultaneously with other users. With each save, everyone editing the workbook at that time can see changes. If you're not ready for others to see your changes, then you can turn off [AutoSave](../../Office-Shared-VBA/articles/about-autosave.md) until you're ready to share your changes and receive others' changes.

**IMPORTANT**: Each instance of add-in or macro code runs independently and maintains its own internal state.

### Example scenario

Imagine an add-in that allows the user to create custom charts based on data in an Excel workbook. This add-in loads data for the user's charts into a hidden sheet in the workbook. When a file containing the custom charts is opened, the add-in reads data on the hidden sheet and loads the chart into memory. As the user makes edits to the chart, this in-memory structure is updated and re-written to the file before each save. This add-in assumes that the only time it is necessary to read the hidden sheet and load it into memory is when the file is opened.

Co-authoring opens another possibility: the hidden sheet could be modified by another user running the same add-in at the same time. If this occurs, the charts that the users are viewing might become out of sync. For example:

- Suppose User A opens the file and starts viewing an existing custom chart.
- While she is doing this, User B opens the same file and starts making changes to the custom chart (for example, changes the zoom level of the chart).
- That change would be saved to the sheet by the add-in on User B’s computer, but User A would never see the change until she reloaded the file.

### Workaround

As much as possible try to avoid making assumptions about when workbook data can be changed. In this case, the developer could modify the add-in to react to the **AfterRemoteChange** event and check the hidden sheet’s values to see if they need to be read again by the add-in to allow User A to view the chart changes that User B made. The macro is intended to be run anytime the chart range is changed. This happens on load and can happen with a remote change. As such, your logic in **AfterRemoteChange** should re-run the macro.

## Integration with co-authoring events

Co-authoring introduces new events **BeforeRemoteChange** and **AfterRemoteChange** which enable you to handle remote changes.

### When to receive remote changes

You may want to receive the latest changes made by another user when the data is fundamental to the expected behavior of the workbook, for example, data visualization and the navigation task pane. 

**Table 1. Examples where the user should receive remote changes**

|**Example**|**Scenario**|
|:-----|:-----|
|Data visualization|Your add-in plots data points on a map based on location data found in a particular range in the workbook. If a user edits any of the location data, all the remote users should receive this change so that each user's map can be updated.|
|Navigation task pane|Your add-in displays all current workbook tabs in a task pane for easy navigation. If a user adds a worksheet, all the remote users should receive this change so that each user's task pane can display a link to the new worksheet.|

### Data visualization example

Let's say that you have created a custom map. In this example, you would add code to change location data then update the map. The workbook is shared with someone in a different city. With autosave on, the change is passed to the remote user who catches the change with the **AfterRemoteChange** event.

```vb
Public Sub UpdateMap()
    'Code that updates map
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Call subroutine that updates map
End Sub
```
```vb
Private Sub Workbook_AfterRemoteChange()
    'Call subroutine that updates map
End Sub
```

**Figure 1. Sample of London map with a few points of interest**
![london locations](images/londonLocations.png) 

### When to ignore remote changes

You may want to avoid changes that cause errors or degraded performance, for example, around data validation and data consistency. See Table 2 for example scenarios.

**Table 2. Examples where the user should not react to remote changes**

|**Example**|**Scenario**|
|:-----|:-----|
|Data validation|A change event is triggered when a specific range is edited in the workbook. Your add-in code then validates the change and, if the check fails, notifies the user via pop-up dialog. However, if all the remote users collaborating on that workbook are notified about a validation failure unrelated to their own changes, this can lead to a poor experience.|
|Data consistency|A change event is triggered and your add-in code synchronizes the data in the workbook with data in another part of the workbook or in an external system. If a remote user receives the change which causes the add-in code to synchronize the same data, this can lead to decreased performance for the remote user or data duplication in the external system.|

### Data validation example

For this example, we've created a chart that displays how much we've made selling various desserts. Neither the cost nor the number of items sold should be negative so there's a validation check that displays a message to the user.  When the invalid value is pushed to the remote users, the validation message should not be displayed to them.

```vb
Public Sub ValidateFigures()
    Dim rangeToValidate As Range
    Set rangeToValidate = ActiveWorkbook.Worksheets("Chart").Range("B2:C6")
    For Each cell In rangeToValidate.Cells
        If (cell.Value < 0) Then
            MsgBox ("Error: Value should not be negative. " & cell.Address)
        End If
    Next
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ActiveWorkbook.Worksheets("Chart").ValidateFigures
End Sub
```

```vb
Private Sub Workbook_AfterRemoteChange()
    ' Do not call validation from RemoteChange event
    'ActiveWorkbook.Worksheets("Chart").ValidateFigures
End Sub
```

**Figure 2. Sample of chart representing desserts sales**
![desserts sales](images/saleschart.png) 

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/about-autosave.md)

#### Additional resources

[Collaborate on Excel workbooks at the same time with co-authoring](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)
