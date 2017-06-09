---
title: Worksheet Object (Excel)
keywords: vbaxl10.chm173072
f1_keywords:
- vbaxl10.chm173072
ms.prod: excel
api_name:
- Excel.Worksheet
ms.assetid: 182b705e-854a-81cc-a4b0-59b942de55ae
ms.date: 06/08/2017
---


# Worksheet Object (Excel)

Represents a worksheet.


## Remarks

The  **Worksheet** object is a member of the **[Worksheets](http://msdn.microsoft.com/library/5ec467a6-97e3-98d7-0b14-845d20c15910%28Office.15%29.aspx)** collection. The **Worksheets** collection contains all the **Worksheet** objects in a workbook.

The  **Worksheet** object is also a member of the [Sheets](http://msdn.microsoft.com/library/048fd93c-bc27-4b58-358f-56fcee1710f8%28Office.15%29.aspx) collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).


## Example

Use  **[Worksheets](http://msdn.microsoft.com/library/8b7d660d-ca49-0bd0-dc57-64defa47bd5e%28Office.15%29.aspx)** (_index_), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.


```
Worksheets(1).Visible = False
```

The worksheet index number denotes the position of the worksheet on the workbook's tab bar.  `Worksheets(1)` is the first (leftmost) worksheet in the workbook, and `Worksheets(Worksheets.Count)` is the last one. All worksheets are included in the index count, even if they're hidden.



The worksheet name is shown on the tab for the worksheet. Use the [Name](http://msdn.microsoft.com/library/3d000cdf-5e81-8701-ca7f-bdcce006363b%28Office.15%29.aspx) property to set or return the worksheet name. The following example protects the scenarios on Sheet1.




```
 
Dim strPassword As String 
strPassword = InputBox ("Enter the password for the worksheet") 
Worksheets("Sheet1").Protect password:=strPassword, scenarios:=True
```

When a worksheet is the active sheet, you can use the  [ActiveSheet](http://msdn.microsoft.com/library/fb5578c3-64a7-edb7-4004-e608739d4c1e%28Office.15%29.aspx) property to refer to it. The following example uses the [Activate](http://msdn.microsoft.com/library/b198dc36-99d0-42db-6cbb-7f68396fd2f5%28Office.15%29.aspx) method to activate Sheet1, sets the page orientation to landscape mode, and then prints the worksheet.




```
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.Orientation = xlLandscape 
ActiveSheet.PrintOut
```

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example uses the BeforeDoubleClick event to open a specified set of files in Notepad. To use this example your worksheet must contain the following data:


- Cell A1 must contain the names of the files to open, each separated by a comma and a space.
    
- Cell D1 must contain the path to where the Notepad files are located.
    
- Cell D2 must contain the path to where the Notepad program is located.
    
- Cell D3 must contain the file extension, without the period, for the Notepad files (txt).
    
When you double-click cell A1, the files specified in cell A1 are opened in Notepad.




```
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
   'Define your variables.
   Dim sFile As String, sPath As String, sTxt As String, sExe As String, sSfx As String
   
   'If you did not double-click on A1, then exit the function.
   If Target.Address <> "$A$1" Then Exit Sub
   
   'If you did double-click on A1, then override the default double-click behaviour with this function.
   Cancel = True
   
   'Set the path to the files, the path to Notepad, the file extension of the files, and the names of the files,
   'based on the information in the worksheet.
   sPath = Range("D1").Value
   sExe = Range("D2").Value
   sSfx = Range("D3").Value
   sFile = Range("A1").Value
   
   'Remove the spaces between the file names.
   sFile = WorksheetFunction.Substitute(sFile, " ", "")
   
   'Go through each file in the list (separated by commas) and
   'create the path, call the executable, and move on to the next comma.
   Do While InStr(sFile, ",")
      sTxt = sPath &amp; "\" &amp; Left(sFile, InStr(sFile, ",") - 1) &amp; "." &amp; sSfx
      If Dir(sTxt) <> "" Then Shell sExe &amp; " " &amp; sTxt, vbNormalFocus
      sFile = Right(sFile, Len(sFile) - InStr(sFile, ","))
   Loop
   
   'Finish off the last file name in the list
   sTxt = sPath &amp; "\" &amp; sFile &amp; "." &amp; sSfx
   If Dir(sTxt) <> "" Then Shell sExe &amp; " " &amp; sTxt, vbNormalNoFocus
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/4fac262c-ea1a-1d2f-bd02-0537c843198c%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/19ea840c-8156-4d9b-8e82-00a687dbc2dc%28Office.15%29.aspx)|
|[BeforeDoubleClick](http://msdn.microsoft.com/library/36e23bc8-0b49-2e22-bfb0-cfff24a82fda%28Office.15%29.aspx)|
|[BeforeRightClick](http://msdn.microsoft.com/library/0263dd09-1648-d3c4-007e-15ef7b82092a%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/c54b75d0-79dd-3e14-0669-447e740e134b%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/d9e11d08-41ba-f0a8-dc55-6c6cd4e76dd0%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/3f66b86b-d0f0-bdc0-594c-3eb9faa44ff2%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/c63eec19-008e-bfb5-1357-3d02426c1bab%28Office.15%29.aspx)|
|[LensGalleryRenderComplete](http://msdn.microsoft.com/library/0e714e01-653b-35ea-455d-21510f59a165%28Office.15%29.aspx)|
|[PivotTableAfterValueChange](http://msdn.microsoft.com/library/097e1c1e-4df6-a0d1-de67-0e0752d2286a%28Office.15%29.aspx)|
|[PivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/220729d9-2da4-53fb-2910-26cc8f835da7%28Office.15%29.aspx)|
|[PivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/4dfcfd60-9249-4eed-1bb3-183b5c567125%28Office.15%29.aspx)|
|[PivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/94a480fa-ce06-e7d7-d4b4-ac21be0525ac%28Office.15%29.aspx)|
|[PivotTableChangeSync](http://msdn.microsoft.com/library/b8cd1e24-4986-d3d4-c37a-b2933c6a9d99%28Office.15%29.aspx)|
|[PivotTableUpdate](http://msdn.microsoft.com/library/66186c97-6855-b360-a6c0-56da617d24a6%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/183e2ca7-06b2-f689-1f77-182dbfbf1e1d%28Office.15%29.aspx)|
|[TableUpdate](http://msdn.microsoft.com/library/69610de6-6884-d5f5-449d-ec1d766d530d%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/b198dc36-99d0-42db-6cbb-7f68396fd2f5%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/7e807ae0-cd97-d95b-f4c4-af1e5674943e%28Office.15%29.aspx)|
|[ChartObjects](http://msdn.microsoft.com/library/234cab0e-a8a2-2174-8881-39b5fb37c743%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/145c7604-5524-b8a2-888c-c3195118cb08%28Office.15%29.aspx)|
|[CircleInvalid](http://msdn.microsoft.com/library/d0e697a4-0c8a-bf2a-06a1-e162696a64dd%28Office.15%29.aspx)|
|[ClearArrows](http://msdn.microsoft.com/library/32b99665-1ac9-9b5d-f009-211a668d6fa6%28Office.15%29.aspx)|
|[ClearCircles](http://msdn.microsoft.com/library/74795226-886b-5922-5448-b93355415bd1%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/ace07575-34f4-a4ae-0922-a2671f2df1ba%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/a51e1673-e09d-824f-1acc-dda18c120204%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/babe18c6-d0ee-62d9-2443-2927cc48a09c%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/e54951d1-6396-c765-7563-1ca7abc16dbd%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/808e6eb8-7811-6f72-5acc-b3779587aa52%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/3f178081-2a42-a751-ae79-8ca149d8ec45%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/65561666-7a47-29d6-2a5d-b5de642a064f%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/8fa41a45-e3d1-29e0-3968-877bcfdf4b57%28Office.15%29.aspx)|
|[PivotTables](http://msdn.microsoft.com/library/b60944cd-827d-15dc-d49e-c739c237de15%28Office.15%29.aspx)|
|[PivotTableWizard](http://msdn.microsoft.com/library/ce37080b-f96f-a706-7b15-7366c268b5cf%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/594f6a4d-29cd-1796-21c2-efc4ed20e067%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/e7065877-2ec9-01ba-4672-4b5a0a8459d2%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/ed517a80-eea9-4268-5fbc-69c659beac0e%28Office.15%29.aspx)|
|[ResetAllPageBreaks](http://msdn.microsoft.com/library/caebf657-3c5b-e465-43e0-88aa3250ba2a%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/2c20ccd0-c4b8-599f-3923-a432caeb6b91%28Office.15%29.aspx)|
|[Scenarios](http://msdn.microsoft.com/library/52e60b55-9316-4c0b-4cb7-ef4605bd31eb%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/2010145e-d36f-7d2b-cfbf-8419c15b31a5%28Office.15%29.aspx)|
|[SetBackgroundPicture](http://msdn.microsoft.com/library/5cff4730-24ba-6147-76c9-e1f9eb970989%28Office.15%29.aspx)|
|[ShowAllData](http://msdn.microsoft.com/library/412acb6c-f83d-44d4-20b5-54a2b7c66284%28Office.15%29.aspx)|
|[ShowDataForm](http://msdn.microsoft.com/library/587a5446-d97e-51d1-d1d9-f5113f8afc0f%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/f955872b-d6bf-5c94-d956-0e84fc7bb9aa%28Office.15%29.aspx)|
|[XmlDataQuery](http://msdn.microsoft.com/library/de728702-962f-a047-a58d-3e2fa9c86acd%28Office.15%29.aspx)|
|[XmlMapQuery](http://msdn.microsoft.com/library/ac1d20f4-92ad-110e-00be-0fe4e098cb35%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/65ca3e7e-2b8f-5882-baaa-17b7658bbf8b%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/766f8501-dae7-32a7-9fae-70a87d0a8eba%28Office.15%29.aspx)|
|[AutoFilterMode](http://msdn.microsoft.com/library/63f33ea5-c9a5-0096-0191-1590cda9d0e1%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/19c14e41-7d8e-b56f-fd60-717df64edee8%28Office.15%29.aspx)|
|[CircularReference](http://msdn.microsoft.com/library/422c447d-a964-c17c-bb43-14254f962a89%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/a734c6d7-3287-3639-6efe-60d270343a44%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/41c18561-2a87-b975-e212-97f39fe10393%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/c2ad8ea7-0fa3-7cde-e3f2-49bbdb81d261%28Office.15%29.aspx)|
|[ConsolidationFunction](http://msdn.microsoft.com/library/4a356e31-723c-88e9-575b-b5a7c5e67757%28Office.15%29.aspx)|
|[ConsolidationOptions](http://msdn.microsoft.com/library/8c09aa4d-86fc-701f-3b89-f8e2be38b948%28Office.15%29.aspx)|
|[ConsolidationSources](http://msdn.microsoft.com/library/d7868b1c-c9ae-97c5-a092-033fe52db5d4%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/39bb2896-2a2f-a7b2-8139-40f0f37104ed%28Office.15%29.aspx)|
|[CustomProperties](http://msdn.microsoft.com/library/49862772-caff-90a1-3266-c8b158003aff%28Office.15%29.aspx)|
|[DisplayPageBreaks](http://msdn.microsoft.com/library/95152278-2618-f200-9933-b6574a49e256%28Office.15%29.aspx)|
|[DisplayRightToLeft](http://msdn.microsoft.com/library/138d361b-d2d0-cf4f-093f-9717dd0f2f6c%28Office.15%29.aspx)|
|[EnableAutoFilter](http://msdn.microsoft.com/library/bff7829a-30f7-3248-e694-ac48621aed31%28Office.15%29.aspx)|
|[EnableCalculation](http://msdn.microsoft.com/library/fc70ae97-b56b-3b57-6f7b-8438b78c424d%28Office.15%29.aspx)|
|[EnableFormatConditionsCalculation](http://msdn.microsoft.com/library/f1f56d9f-3a0f-e3d4-f686-1a695a55604e%28Office.15%29.aspx)|
|[EnableOutlining](http://msdn.microsoft.com/library/db849ddf-871d-19cd-9765-3194a8c1e34e%28Office.15%29.aspx)|
|[EnablePivotTable](http://msdn.microsoft.com/library/8cd09896-9752-677f-a7fd-da46d68ac42a%28Office.15%29.aspx)|
|[EnableSelection](http://msdn.microsoft.com/library/e1647c07-3863-9268-864c-1c62b2eebbb1%28Office.15%29.aspx)|
|[FilterMode](http://msdn.microsoft.com/library/d9bcaa8a-caf3-96a4-445d-d957a987b057%28Office.15%29.aspx)|
|[HPageBreaks](http://msdn.microsoft.com/library/0d26aa71-714f-a6a0-8a10-4ea6bd7d852d%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/ac2fe50a-23a0-9982-d448-b18a91092624%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/970065b3-f9bd-d518-261a-f5f704c350df%28Office.15%29.aspx)|
|[ListObjects](http://msdn.microsoft.com/library/29c20c8d-aa64-f578-2c8a-5567651ba44c%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/9490f86c-a82f-d1ab-7315-29b89c799301%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/3d000cdf-5e81-8701-ca7f-bdcce006363b%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/4bdccfa9-7aa1-c3d6-6a89-5ce24aad2ad2%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/971d5df0-ba23-ac67-7862-67586452e992%28Office.15%29.aspx)|
|[Outline](http://msdn.microsoft.com/library/e53d8038-f20b-9d55-1ee0-c5f6b4a099d4%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/50310e49-49d0-c8c6-ed23-0eacae6355b5%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/01ab7232-7b94-fc4f-9fe1-e5592d8b9ee6%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/8409e3c6-564e-2ba1-1e49-79a1c37cc845%28Office.15%29.aspx)|
|[PrintedCommentPages](http://msdn.microsoft.com/library/3ade9c86-c6b9-08fa-3bc6-a040dd1da36a%28Office.15%29.aspx)|
|[ProtectContents](http://msdn.microsoft.com/library/807717f6-1265-2d5d-5221-bc46b24d8281%28Office.15%29.aspx)|
|[ProtectDrawingObjects](http://msdn.microsoft.com/library/a3733b3b-dca4-4131-e197-5c919d44c7bd%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/46bf2025-46cf-81ae-4864-2d6266dab173%28Office.15%29.aspx)|
|[ProtectionMode](http://msdn.microsoft.com/library/465e2405-c9f3-83ac-f68d-ff9172375e1f%28Office.15%29.aspx)|
|[ProtectScenarios](http://msdn.microsoft.com/library/7b0aacea-00f3-7f0a-2be1-693f0efbec88%28Office.15%29.aspx)|
|[QueryTables](http://msdn.microsoft.com/library/1228c6e0-f8d9-87a3-2fbf-1526f5229f1b%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/9a323305-c822-ef9e-1cc8-ec077a976834%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/5d07304e-a3c9-2a75-b2ba-4a7b16ce6516%28Office.15%29.aspx)|
|[ScrollArea](http://msdn.microsoft.com/library/7421676d-3a98-3826-31f9-80e7c8946777%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/6206b5e8-742d-797f-12ee-daf3225a53dc%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/2e5cadb0-a688-5864-5974-861270b61db3%28Office.15%29.aspx)|
|[StandardHeight](http://msdn.microsoft.com/library/43dd7321-5df7-2cda-9e51-75f81ab203f2%28Office.15%29.aspx)|
|[StandardWidth](http://msdn.microsoft.com/library/6792ce79-0a73-fcbd-ea52-7d7aee7b9932%28Office.15%29.aspx)|
|[Tab](http://msdn.microsoft.com/library/386edcb0-868e-3f24-b4f0-8e52b9fcffcb%28Office.15%29.aspx)|
|[TransitionExpEval](http://msdn.microsoft.com/library/a92d8efb-5f19-4b41-11b2-a20b3ad5bf1d%28Office.15%29.aspx)|
|[TransitionFormEntry](http://msdn.microsoft.com/library/ec17c4db-d94e-2fd9-39eb-7c1e0ea40a49%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/af99d12c-eddf-f649-d70c-6ad5efc0920f%28Office.15%29.aspx)|
|[UsedRange](http://msdn.microsoft.com/library/f004b93c-d785-de19-1fb4-bbe0b2e9b6cd%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/48860564-6079-932e-2cae-0802235be61e%28Office.15%29.aspx)|
|[VPageBreaks](http://msdn.microsoft.com/library/2a8d5c77-a609-4995-7216-de71295eda9a%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Other resources

[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

