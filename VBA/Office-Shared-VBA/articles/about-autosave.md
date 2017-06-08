---
title: About AutoSave in Office
ms.prod: MULTIPLEPRODUCTS
ms.assetid: about-autosave
ms.date: 06/08/2017
---


# About AutoSave in Office

Learn about how AutoSave works in Excel, PowerPoint, and Word and how it can impact add-in or macro integration. For more about how AutoSave works in general, see [this link](https://support.office.com/en-US/article/What-is-AutoSave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5).

## Overview of AutoSave

When a file is hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), AutoSave enables the user's edits to be saved automatically and continuously. When the file is shared with others, then the others' changes will be merged into this user's version of the file. If AutoSave is turned off, then save must be triggered manually for the user's changes to be persisted in the cloud and for that user to receive others' changes.

Currently, Excel, Word, and PowerPoint provide a **BeforeSave** event that allows a developer to execute code after the user triggers a save but before the save actually occurs. Excel also provides an **AfterSave** event that can execute macro or add-in code after the save completes.

When AutoSave is enabled, these events will fire automatically on a periodic basis without user interaction. Because of this new behavior when AutoSave is enabled, add-ins and macros that leverage these events may experience problems when AutoSave is on.

In general, these issues can be avoided if the user chooses to disable AutoSave. You can also do this on the user’s behalf using the **autoSaveOn** property (see the example below). You can also take steps as a developer to mitigate these potential problems so that your add-ins and macros work smoothly, even if AutoSave is enabled.

## Potential issues with save events and AutoSave

### Issue 1: Code in **BeforeSave** or **AfterSave** events runs too long

#### Description

In general, Word, Excel and PowerPoint are not responsive to user interaction while add-in or macro code is being run. Therefore, if your code in a BeforeSave or AfterSave event handler takes too long to run, it may significantly degrade the user experience. When AutoSave is disabled, this code is only run when the user explicitly chooses to save, so a delay is not as noticeable and can be avoided by the user until he or she is ready to save. When AutoSave is enabled, this code will run automatically on a periodic basis, which has the potential to interrupt the user, especially if the code takes a long time.

#### Example Scenario

Imagine an add-in that allows the user to create custom maps based on data in an Excel workbook. Such an add-in might have **BeforeSave** code that serializes any maps that the user has created and stores them in the workbook in a CustomXML part. This process might take a second to complete and Excel is unresponsive while this is happening. When AutoSave is off, the user gets to choose when he or she wants to save and therefore, even though the add-in slows down the save process slightly, the user does not notice. When AutoSave is enabled, this **BeforeSave** code is executed automatically on a periodic basis even if the user is in the middle of something else (like typing data into a cell) which could be extremely annoying.

#### Workarounds

- Add-ins should try to avoid running long-running operations inside of a save event. In this example, the developer could choose to persist the custom maps to the file as they are created or modified by the user rather than waiting for the save event.

### Issue 2: Code in save events surfaces a modal dialog

#### Description

Any code that runs in a save event that displays UI such as a modal dialog has the potential to seriously degrade the user experience when AutoSave is on. Because the **BeforeSave** and **AfterSave** events run automatically on a periodic basis, these dialog boxes may interrupt the user's normal workflow.

#### Example Scenario

An add-in that validates a Word document before save to ensure that company branding is applied might fire a dialog box that alerts the user about any problems that were found and offers a way to correct them. Since the **BeforeSave** event now fires automatically and continuously, this validation dialog might appear suddenly while the user is doing something else.

#### Workarounds

- Consider removing any code that needs to display UI to other areas of the application. For example, there could be a "validate" button that the user could click to trigger the validation process. Or you could fire the validation code only if the user attempts to change the existing data.

- If you want validation code to trigger only on the first save from a new document but not on subsequent auto-saves, consider inspecting a property like Excel's **Workbook.Path** before displaying any UI during **BeforeSave** or **AfterSave**. In Excel, the **Workbook.Path** property should be blank if the workbook does not yet have a save location.

### Issue 3: Code in save events causes the undo stack to get blown away (Excel Only)

#### Description

In general, executing certain VBA statements in Excel will clear the undo stack. For example, changing the value of a cell by executing
`ActiveCell.Value = "asdf"` will clear the undo stack. If such code is present in the **BeforeSave** or **AfterSave** event for a particular macro or add-in and AutoSave is on, a user of that macro or add-in will frequently not be able to undo normal user actions as expected.

#### Example Scenario

An add-in might have code that runs in response to the **BeforeSave** event that inspects the document and writes values to a "log" table in the workbook. When AutoSave is on, this would clear the undo stack periodically, which can potentially annoy users.

#### Workarounds

- Consider removing code that writes to the workbook in **BeforeSave** or **AfterSave** events. For example, the add-in described above might be modified to store the change log in a separate file or database.

### Issue 4: Code that dirties the workbook in **AfterSave** (Excel Only)

#### Description

When AutoSave is on, the **BeforeSave** and **AfterSave** events will only trigger if there has been a change in the workbook since the last time they were triggered. If code in the **AfterSave** event dirties the workbook (i.e. makes additional changes), that could potentially trigger events to fire again after a while for the same change then queue up the events to fire again, and so on, indefinitely. This could waste system resources and battery life.

#### Workarounds

- Code that dirties the workbook in **AfterSave** should be moved to **BeforeSave** or to another location entirely (see Issue 3 above). This isn't good practice today, even without AutoSave, since it leaves the workbook in a perpetual "dirty" state, which causes a prompt to appear on close that asks the user to save their changes even if they made no additional changes. 

### Issue 5: Code that cancels the file save in **BeforeSave** (by setting Cancel argument to True)

#### Description

Today, it is possible to cancel the save in the **BeforeSave** event by setting cancel to True: 
```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean) 
    Cancel = True 
End Sub
```
When AutoSave is enabled, the application (that is, Excel, Word, or PowerPoint) will trigger saves automatically on a continuous basis until the file has no more unsaved changes. After a single change is made to the file by the user, the application will attempt to save it. If the developer chooses to cancel the save in the manner described above, the application continually determines that there is are unsaved changes, which will cause the save to (eventually) be attempted again. Since the same event code that cancelled the first save will also cancel this second save attempt, the process will continue for as long as the file is open, potentially degrading performance and battery life.

#### Example Scenario

An add-in might completely override the default Word save code so that the file is actually saved to a corporate database instead of a disk or sharepoint location. Such an add-in would first cancel the attempted save before trying to save in another place.

#### Workarounds

- Such add-ins should ensure that AutoSave is turned off by setting autoSaveOn to False. Since a file must already be saved in a OneDrive or SharePoint location for AutoSave to be on, AutoSave should already be off in most cases in the scenario above.

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

[Co-authoring](../../Excel-VBA/articles/about-coauthoring-in-excel.md)

[Workbook Object](../../Excel-VBA/articles/workbook-object-excel.md)

#### Additional resources

[What is AutoSave?](https://support.office.com/en-US/article/What-is-AutoSave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5)