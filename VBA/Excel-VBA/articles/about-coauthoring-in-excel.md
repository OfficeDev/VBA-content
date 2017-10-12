---
title: About coauthoring in Excel
ms.prod: excel
ms.date: 07/19/2017
---


# About coauthoring in Excel

Learn about how coauthoring works in Excel and how you may need to adjust your add-in or macro for smooth integration with coauthoring.

Coauthoring is available to all Excel Online users. This feature is also available on Excel for Windows Desktop, but only to Office 365 customers.

## Introduction to coauthoring

[Coauthoring](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104) allows you to edit a workbook hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online) simultaneously with other users. With each save, everyone editing the workbook at that time can see changes. With [AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md) enabled, you can see everyone's changes to the workbook in real-time. If you're not ready for others to see your changes, you can turn off AutoSave until you're ready to share your changes and receive others' changes.

## Principles of coauthoring

Excel will automatically synchronize changes that are made to the workbook (whether by a user or your code). For example, let's say that code is running on a user's instance and modifies the contents of a cell like this: `Range("A1").Value = "myNewValue"`. Excel would take care of sending this change to other coauthors. 

Now let's say there's code running on another user's instance that inspects the contents of that cell like this: `MsgBox Range("A1").Value`. The second user would see the value `"myNewValue"` that had been set by the first user.

However, Excel will *not* automatically synchronize any variables created by your code outside of the workbook content. For example, let's say that your code reads a value from a cell, and then loads it into a variable:

```vb
Dim myVariable
myVariable = Range("A1").Value
```

Excel will not automatically update the value of `myVariable`, meaning that `myVariable` will not be kept in sync with a variable of the same name that's created by code running on the other coauthors' Excel instances.

## Situations where you may need to adapt your solution to a coauthoring environment

Because existing add-ins and macros can rely on Excel to seamlessly transmit the changes they make to the workbook to the coauthors, you can usually use your code in this new environment without making any changes or updates. However, in two cases, you may need to adapt your code if you want it to work smoothly in a coauthoring setting:

- [Add-ins that have an internal, in-memory state outside of the workbook content](#StateOfAddins)
- [Add-ins that leverage events](#UseEvents)

[**BeforeRemoteChange**](workbook-beforeremotechange-event-excel.md) and [**AfterRemoteChange**](workbook-afterremotechange-event-excel.md) events were added to enable you to manage remote changes where applicable.

### <a name="StateOfAddins"></a>Add-ins that have an internal, in-memory state outside of the workbook content

Imagine an add-in that allows the user to create custom charts based on data in an Excel workbook. This add-in loads data for the user's charts into a hidden sheet in the workbook. When a user opens a file containing the custom charts, the add-in reads data on the hidden sheet and loads the chart into memory. As the user edits the chart, this in-memory structure is updated and re-written to the file before each save. This add-in assumes that the only time it is necessary to read the hidden sheet and load it into memory is when the file is opened. 

Coauthoring opens another possibility: the hidden sheet could be modified by another user running the same add-in at the same time. If this occurs, the charts that the users are viewing might become out of sync. For example:
- Suppose User A opens the file and starts viewing an existing custom chart.
- While she is doing this, User B opens the same file and starts making changes to the custom chart (for example, changes the type of chart).
- That change would be saved to the sheet by the add-in on User B’s computer, but User A would never see the change until she reloaded the file.

#### Workaround 

As much as possible, try to avoid making assumptions about when workbook data can be changed. In this case, you could modify the add-in to react to the **AfterRemoteChange** event, and then check the hidden sheet’s values to see if they need to be read again by the add-in to allow User A to view the chart changes that User B made. The add-in is intended to be run anytime the chart range is changed. This happens on load and can happen with a remote change. As such, your logic in **AfterRemoteChange** should re-run the add-in. 

### <a name="UseEvents"></a>Add-ins that leverage events

Your add-in or macro may already subscribe to save or change events. With the introduction of coauthoring, you may experience issues with:

- [**BeforeSave** or **AfterSave** events](#SaveEvents)
- [Change events](#ChangeEvents)

#### <a name="SaveEvents"></a>Save events

You may experience issues when your code uses save events such as **BeforeSave** and **AfterSave**. For more information, see [Potential issues with save events and AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md#IssuesWithSaveEventsAndAutoSave).

#### <a name="ChangeEvents"></a>Change events

By default, your code usually does not need to handle changes from remote users. However, there are some cases where handling remote changes may cause problems. Two sample scenarios are explored here.

#### Sample scenario: Data validation

A change event is triggered when a specific range is edited in the workbook. Your add-in code then validates the change and, if the check fails, notifies the user via a pop-up window. However, if all the remote users collaborating on that workbook are notified about a validation failure unrelated to their own changes, this can lead to a poor experience.

#### Example

For this example, a chart was created that displays how much was made selling various desserts. Neither the cost nor the number of items sold should be negative, so there's a validation check that displays a message to the user.  When the invalid value is pushed to the remote users, the validation message should not be displayed to them.

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

<br/>

As such, there is no need to subscribe to either the **BeforeRemoteChange** or **AfterRemoteChange** event in this case.

```vb
Private Sub Workbook_AfterRemoteChange()
    ' Do not call validation from RemoteChange event
    'ActiveWorkbook.Worksheets("Chart").ValidateFigures
End Sub
```

<br/>

*Figure 1. Sample of chart representing dessert sales*

![dessert sales](images/saleschart.png) 

<br/>

#### Sample scenario: Data consistency

A change event is triggered, and your add-in code synchronizes the data in the workbook with data in another part of the workbook or in an external system. If a remote user receives the change that causes the add-in code to synchronize the same data, this can lead to decreased performance for the remote user or data duplication in the external system.

### Potential issues with change events

Although normally you would not want your event handler code to run in response to changes from a remote user, the default behavior of *not* firing change events could cause problems. Following are some examples of problems and how you can work around them by using **BeforeRemoteChange** and **AfterRemoteChange** events.

#### Sample scenario: Data visualization

Your add-in plots data points on a map based on location data found in a range in the workbook. If a user edits any of the location data, all the remote users should receive this change so that each user's map can be updated.

#### Example

Let's say that you have created a custom map. In this example, you could add code to change location data, and then update the map. The workbook is shared with someone in a different city. With AutoSave on, the change is passed to the remote user, but that user's map will not be updated.

```vb
Public Sub UpdateMap()
    'Code that updates map
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Call subroutine that updates map
End Sub
```

<br/>

Now use the **AfterRemoteChange** event to add code that updates the map. Subsequent changes sent to the remote user are used to update the map.

```vb
Private Sub Workbook_AfterRemoteChange()
    'Call subroutine that updates map
End Sub
```

<br/>

*Figure 2. Sample of London map with a few points of interest*

![london locations](images/londonLocations.png) 

<br/>

#### Sample scenario: Navigation task pane

Your add-in displays all current workbook tabs in a task pane for easy navigation. If a user adds a worksheet, all the remote users should receive this change so that each user's task pane can display a link to the new worksheet.

## See also

#### Concepts

[AutoSave](../../Office-Shared-VBA/articles/how-autosave-impacts-addins-and-macros.md)

#### Additional resources

[Collaborate on Excel workbooks at the same time with coauthoring](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104)
