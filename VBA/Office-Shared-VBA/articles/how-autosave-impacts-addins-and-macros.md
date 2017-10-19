---
title: How AutoSave impacts add-ins and macros
ms.prod: office
ms.date: 07/28/2017
---


# How AutoSave impacts add-ins and macros

Learn about how AutoSave works in Excel, PowerPoint, and Word, and how it can impact add-ins or macros. For more information about how AutoSave works in general, see ["What is AutoSave?"][AutoSaveArticle].

## Overview of AutoSave

When a file is hosted in the cloud (that is, OneDrive, OneDrive for Business, or SharePoint Online), AutoSave enables the user's edits to be saved automatically and continuously. When the file is shared with others, their changes are merged into this user's version of the file. If AutoSave is turned off, save must be triggered manually for the user's changes to be persisted in the cloud and for this user to receive others' changes.

Currently, Excel, Word, and PowerPoint provide a **BeforeSave** event that allows a developer to execute code after the user triggers a save but before the save occurs. Excel also provides an **AfterSave** event that can execute macro or add-in code after the save completes.

When AutoSave is enabled, these events fire automatically on a periodic basis without user interaction. Because of this, add-ins and macros that leverage these events may experience problems when AutoSave is on.

In general, these issues can be avoided if the user chooses to disable AutoSave. You can do this on the user’s behalf by using the **AutoSaveOn** property in [Word][AutoSaveOn_Word], [Excel][AutoSaveOn_Excel], and [PowerPoint][AutoSaveOn_PowerPoint] if it's available (see the following example). You can also take steps as a developer to mitigate these problems so that your add-ins and macros work smoothly, even if AutoSave is enabled.

### <a name="example"></a>Example

This example turns off AutoSave and notifies the user that the workbook is not being automatically saved.

```vb
Sub UseAutoSaveOn()
    ActiveWorkbook.AutoSaveOn = False
    MsgBox "This workbook is being saved automatically: " & ActiveWorkbook.AutoSaveOn
End Sub
```

<br/>

## <a name="IssuesWithSaveEventsAndAutoSave"></a>Potential issues with save events and AutoSave

You may need to handle one or more of the following issues regarding the interaction between save events and AutoSave:

1. Code in **BeforeSave** or **AfterSave** events runs too long
2. Code in save events surfaces a modal dialog
3. Code in save events clears the undo stack (Excel only)
4. Code in **AfterSave** dirties the workbook (Excel only)
5. Code in **BeforeSave** cancels the file save (by setting Cancel argument to True)

<br/>

### <a name="Issue1"></a>Issue 1: Code in BeforeSave or AfterSave events runs too long

In general, Word, Excel and PowerPoint are not responsive to user interaction while add-in or macro code is being run. Therefore, if your code in a **BeforeSave** or **AfterSave** event handler takes too long to run, it may significantly degrade the user experience. 

When AutoSave is disabled, this code is only run when the user explicitly chooses to save, so a delay is not as noticeable and can be avoided by the user until he or she is ready to save. 

When AutoSave is enabled, this code runs automatically on a periodic basis, which has the potential to interrupt the user, especially if the code takes a long time.

#### Example scenario

Imagine an add-in that allows the user to create custom maps based on data in an Excel workbook. Such an add-in might have **BeforeSave** code that serializes any maps that the user has created and stores them in the workbook in a CustomXML part. This process might take a second to complete, and Excel could be unresponsive while this is happening. 

When AutoSave is off, the user gets to choose when he or she wants to save, and therefore, even though the add-in slows down the save process slightly, the user does not notice. 

When AutoSave is enabled, this **BeforeSave** code runs automatically on a periodic basis even if the user is in the middle of something else (such as typing data into a cell), which could be extremely annoying.

#### Workaround

Add-ins should try to avoid long-running operations inside of a save event. In this example, the developer could choose to persist the custom maps to the file as they are created or modified by the user, rather than waiting for the save event.

<br/>

### <a name="Issue2"></a>Issue 2: Code in save events surfaces a modal dialog

Any code that runs in a save event that displays UI such as a modal dialog has the potential to seriously degrade the user experience when AutoSave is on. Because the **BeforeSave** and **AfterSave** events run automatically on a periodic basis, these dialog boxes may interrupt the user's normal workflow.

#### Example scenario

An add-in that validates a Word document before save to ensure that company branding is applied might fire a dialog box that alerts the user about any problems that were found and offers a way to correct them. Because the **BeforeSave** event now fires automatically and continuously, this validation dialog might appear suddenly while the user is doing something else.

#### Workarounds

Consider removing any code that needs to display UI to other areas of the application. For example, the user could click a "validate" button to trigger the validation process, or you could fire the validation code only if the user attempts to change the existing data.

If you want validation code to trigger only on the first save from a new document but not on subsequent auto-saves, consider inspecting a property such as Excel's **Workbook.Path** before displaying any UI during **BeforeSave** or **AfterSave**. In Excel, the **Workbook.Path** property should be blank if the workbook does not yet have a save location.

<br/>

### <a name="Issue3"></a>Issue 3: Code in save events clears the undo stack (Excel only)

In general, if you run certain VBA statements in Excel, the undo stack will be cleared. For example, if you change the value of a cell by running `ActiveCell.Value = "myValue"`, the undo stack is cleared. If such code is present in the **BeforeSave** or **AfterSave** event for a macro or add-in, and AutoSave is on, a user of that macro or add-in will frequently not be able to undo normal user actions as expected.

#### Example scenario

An add-in might have code that runs in response to the **BeforeSave** event that inspects the document and writes values to a "log" table in the workbook. When AutoSave is on, this would clear the undo stack periodically, which can potentially annoy users.

#### Workaround

Consider removing code that writes to the workbook in **BeforeSave** or **AfterSave** events. For example, the add-in described in the example scenario might be modified to store the change log in a separate file or database.

<br/>

### <a name="Issue4"></a>Issue 4: Code in **AfterSave** dirties the workbook (Excel only)

When AutoSave is on, the **BeforeSave** and **AfterSave** events will only trigger if there has been a change in the workbook since the last time they were triggered. If code in the **AfterSave** event dirties the workbook (that is, makes additional changes), that could potentially trigger events to fire again for the same change, and then queue up the events to fire again indefinitely. This could waste system resources and affect battery life.

#### Workaround

Code that dirties the workbook in **AfterSave** should be moved to **BeforeSave** or to another location entirely (see [Issue 3](#Issue3)). This isn't a good practice today, even without AutoSave, because it leaves the workbook in a perpetual "dirty" state, which causes a prompt to appear on close that asks the user to save their changes even if they made no additional changes. 

<br/>

### <a name="Issue5"></a>Issue 5: Code in **BeforeSave** cancels the file save (by setting Cancel argument to True)

Today, it is possible to cancel the save in the **BeforeSave** event by setting `Cancel` to True:

```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean) 
    Cancel = True
End Sub
```

<br/>

When AutoSave is enabled, the application (that is, Excel, Word, or PowerPoint) triggers saves automatically on a continuous basis until the file has no more unsaved changes. After the user makes a single change to the file, the application attempts to save it. 

If the developer chooses to cancel the save in the manner described earlier, the application continually determines that there are unsaved changes, which causes the save to (eventually) be attempted again. Because the same event code that cancelled the first save will also cancel this second save attempt, the process will continue for as long as the file is open, potentially degrading performance and battery life.

#### Example scenario

An add-in might completely override the default Word save code so that the file is saved to a corporate database instead of to a disk or SharePoint location. Such an add-in would first cancel the attempted save before trying to save in another place.

#### Workaround

Such add-ins should ensure that AutoSave is turned off by setting AutoSaveOn to False. Because a file must already be saved in a OneDrive or SharePoint location for AutoSave to be on, AutoSave should already be off or disabled in most versions of this scenario.

## See also

#### Concepts

- [Coauthoring in Excel](../../Excel-VBA/articles/about-coauthoring-in-excel.md)

- [Document object](../../Word-VBA/articles/document-object-word.md)

- [Presentation object](../../PowerPoint-VBA/articles/presentation-object-powerpoint.md)

- [Workbook object](../../Excel-VBA/articles/workbook-object-excel.md)

#### Additional resources

- [What is AutoSave?][AutoSaveArticle]

- [**AutoSaveOn** property in Excel][AutoSaveOn_Excel]

- [**AfterSave** event in Excel](../../Excel-VBA/articles/application-workbookaftersave-event-excel.md)

- [**BeforeSave** event in Excel](../../Excel-VBA/articles/application-workbookbeforesave-event-excel.md)

- [**AutoSaveOn** property in PowerPoint][AutoSaveOn_PowerPoint]

- [**BeforeSave** event in PowerPoint](../../PowerPoint-VBA/articles/application-presentationbeforesave-event-powerpoint.md)

- [**AutoSaveOn** property in Word][AutoSaveOn_Word]

- [**BeforeSave** event in Word](../..//Word-VBA/articles/application-documentbeforesave-event-word.md)

[comment]: # (Local shared links)

[AutoSaveArticle]: https://support.office.com/en-US/article/What-is-AutoSave-6d6bd723-ebfd-4e40-b5f6-ae6e8088f7a5

[AutoSaveOn_Excel]: ../../Excel-VBA/articles/workbook-autosaveon-property-excel.md

[AutoSaveOn_PowerPoint]: ../../PowerPoint-VBA/articles/presentation-autosaveon-property-powerpoint.md

[AutoSaveOn_Word]: ../../Word-VBA/articles/document-autosaveon-property-word.md
