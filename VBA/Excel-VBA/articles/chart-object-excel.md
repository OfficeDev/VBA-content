---
title: Chart Object (Excel)
keywords: vbaxl10.chm147072
f1_keywords:
- vbaxl10.chm147072
ms.prod: excel
api_name:
- Excel.Chart
ms.assetid: 179c32ce-49bd-6f36-ea12-89fb5443f3ea
ms.date: 06/08/2017
---


# Chart Object (Excel)

Represents a chart in a workbook.


## Remarks

The chart can be either an embedded chart (contained in a **[ChartObject](http://msdn.microsoft.com/library/b546e6f2-7ac6-2dea-eba2-f98f68f3df65%28Office.15%29.aspx)** object) or a separate chart sheet.

The following properties and methods for returning a **Chart** object are described in the example section:

-  **Charts** method
    
-  **ActiveChart** property
    
-  **ActiveSheet** property
    

## Example

The **[Charts](http://msdn.microsoft.com/library/06d4602e-a713-7ca0-db39-2d8a29f084a0%28Office.15%29.aspx)** collection contains a **Chart** object for each chart sheet in a workbook. Use **Charts** ( _index_ ), where index is the chart-sheet index number or name, to return a single **Chart** object. 

The chart index number represents the position of the chart sheet on the workbook tab bar. _Charts(1)_ is the first (leftmost) chart in the workbook; _Charts(Charts.Count)_ is the last (rightmost). 

All chart sheets are included in the index count, even if they are hidden. The chart-sheet name is shown on the workbook tab for the chart. You can use the **[Name](http://msdn.microsoft.com/library/3da85312-f508-499a-6799-c1e15e2259a0%28Office.15%29.aspx)** property to set or return the chart name. 

The following example changes the color of series 1 on chart sheet 1.

```
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

The following example moves the chart named Sales to the end of the active workbook.

```
Charts("Sales").Move after:=Sheets(Sheets.Count)
```

The **Chart** object is also a member of the **[Sheets](http://msdn.microsoft.com/library/048fd93c-bc27-4b58-358f-56fcee1710f8%28Office.15%29.aspx)** collection, which contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** ( _index_ ), where _index_ is the sheet index number or name, to return a single sheet.

When a chart is the active object, you can use the **ActiveChart** property to refer to it. A chart sheet is active if the user has selected it or if it has been activated with the **[Activate](http://msdn.microsoft.com/library/b2bda196-4f0c-252f-cd6f-79c9f3d08f7c%28Office.15%29.aspx)** method of the **Chart** object or the **[Activate](http://msdn.microsoft.com/library/21997b8b-e446-249b-b33e-ee3b7f9aa564%28Office.15%29.aspx)** method of the **ChartObject** object. 

The following example activates chart sheet 1 and then sets the chart type and title.

```
Charts(1).Activate 
With ActiveChart 
 .Type = xlLine 
 .HasTitle = True 
 .ChartTitle.Text = "January Sales" 
End With
```

An embedded chart is active if the user has selected it, or the **[ChartObject](http://msdn.microsoft.com/library/b546e6f2-7ac6-2dea-eba2-f98f68f3df65%28Office.15%29.aspx)** object in which it is contained has been activated with the **[Activate](http://msdn.microsoft.com/library/21997b8b-e446-249b-b33e-ee3b7f9aa564%28Office.15%29.aspx)** method. 

The following example activates embedded chart 1 on worksheet 1 and then sets the chart type and title. Notice that after the embedded chart has been activated, the code in this example is the same as that in the previous example. Using the **ActiveChart** property allows you to write Visual Basic code that can refer to either an embedded chart or a chart sheet (whichever is active).

```
Worksheets(1).ChartObjects(1).Activate 
ActiveChart.ChartType = xlLine 
ActiveChart.HasTitle = True 
ActiveChart.ChartTitle.Text = "January Sales"
```

When a chart sheet is the active sheet, you can use the **ActiveSheet** property to refer to it. The following example uses the **Activate** method to activate the chart sheet named Chart1, and then sets the interior color for series 1 in the chart to blue.

```
Charts("chart1").Activate 
ActiveSheet.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
```

## Events

|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/7b878d1b-3059-93cb-389a-a2633f613a4d%28Office.15%29.aspx)|
|[BeforeDoubleClick](http://msdn.microsoft.com/library/406c6b9f-1182-5f5b-b954-afe10cd21a9b%28Office.15%29.aspx)|
|[BeforeRightClick](http://msdn.microsoft.com/library/d01f6911-2f6b-3118-27a2-dfafa48791ab%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/5510a6e9-5038-9bd2-8f7b-aa75427f48d4%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/b843b64a-ad20-d160-1abb-88317114b44c%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/6c4ef5ce-560e-a7d5-c602-99a999fb5535%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/b1277953-a882-f00f-2ac1-dd0cc49fef72%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/45281aac-a4f6-390d-e767-a4fe2ee670fc%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/d1b7d0bb-d190-18f2-83f9-b91b637d80aa%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/00ea6501-e92e-5b95-f2b0-bb9b014bb5ec%28Office.15%29.aspx)|
|[SeriesChange](http://msdn.microsoft.com/library/80a8058c-0445-0051-24d1-1a965c302790%28Office.15%29.aspx)|

## Methods

|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/b2bda196-4f0c-252f-cd6f-79c9f3d08f7c%28Office.15%29.aspx)|
|[ApplyChartTemplate](http://msdn.microsoft.com/library/b4695f3f-26ac-1e35-7318-0091d9b1f130%28Office.15%29.aspx)|
|[ApplyDataLabels](http://msdn.microsoft.com/library/20966609-9713-c644-81d7-196b06169975%28Office.15%29.aspx)|
|[ApplyLayout](http://msdn.microsoft.com/library/0e07936d-c179-9b38-a6d4-1d71d1c5af3b%28Office.15%29.aspx)|
|[Axes](http://msdn.microsoft.com/library/d0520f61-9aff-894b-9975-37dcb5b5fe3c%28Office.15%29.aspx)|
|[ChartGroups](http://msdn.microsoft.com/library/dffa4fc3-b2db-eb50-b309-95e99972525f%28Office.15%29.aspx)|
|[ChartObjects](http://msdn.microsoft.com/library/5b518ecf-9c1a-fb2f-c833-182c37b8c2c1%28Office.15%29.aspx)|
|[ChartWizard](http://msdn.microsoft.com/library/c47588d9-6969-d6bb-cbbc-4941198d78b4%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/9c39b0f1-4401-1399-58fa-444c9fa9fab4%28Office.15%29.aspx)|
|[ClearToMatchColorStyle](http://msdn.microsoft.com/library/5b409cca-e458-21dd-77cc-0a93df1d4539%28Office.15%29.aspx)|
|[ClearToMatchStyle](http://msdn.microsoft.com/library/8e45ac2f-c479-30b2-c0b0-3c1cf0670a80%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/c7294fd6-286a-774d-9dd8-4db33a59b10f%28Office.15%29.aspx)|
|[CopyPicture](http://msdn.microsoft.com/library/f69451cd-4be5-982a-58b8-63e0f24e0261%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/700df0f8-8d85-d8dc-aaa6-c72dcd4a0277%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/7a171fd5-e084-7172-f429-5425e0d342d4%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/4dc7dea6-9be8-ccd4-8198-7726b8fad024%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/4fa48315-c9e9-944c-71c3-72ec9894daac%28Office.15%29.aspx)|
|[FullSeriesCollection](http://msdn.microsoft.com/library/875c18cf-064f-6b2f-2650-f5d07c16bc4d%28Office.15%29.aspx)|
|[GetChartElement](http://msdn.microsoft.com/library/a4888d1b-f73b-43cd-5318-95c1d63944fa%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/3744f7f3-f7df-3ac2-48b7-b57ce3a8c812%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/ec8c8eae-17a8-20a0-a87c-81f31b21d735%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/e42150c1-8661-75b4-f1e8-fec8cc82f59b%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/e34d3d30-39f8-dbd4-1a39-d3ef9f84e0f4%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/6529b0d5-5347-fcbc-f12a-3ab9e8c01359%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/c08ad230-8bec-efd0-b94a-92b2324b5925%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/5f46d721-021b-d615-12c6-78aab49df500%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/4ede937c-d710-521d-dfeb-0af21ee6ba7d%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/21e2a786-1df2-21ea-f32f-81e07dc2261c%28Office.15%29.aspx)|
|[SaveChartTemplate](http://msdn.microsoft.com/library/d9e36023-b5bb-aaf4-5b34-9a22df468ced%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/20f866f4-14b9-075c-372c-47a9f536f0c3%28Office.15%29.aspx)|
|[SeriesCollection](http://msdn.microsoft.com/library/0a628f00-1ee6-9ff8-dce1-c7aabbdd1a85%28Office.15%29.aspx)|
|[SetBackgroundPicture](http://msdn.microsoft.com/library/11a2d89d-d568-b30f-7f8c-e56495879ac4%28Office.15%29.aspx)|
|[SetDefaultChart](http://msdn.microsoft.com/library/8be43de3-8b7d-4885-3e49-19aa0c65564f%28Office.15%29.aspx)|
|[SetElement](http://msdn.microsoft.com/library/0efff437-179b-fe16-118b-6f3cde49c5cf%28Office.15%29.aspx)|
|[SetSourceData](http://msdn.microsoft.com/library/fc41cc05-087a-f53c-2f54-fd6307de51d6%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/59a367bd-037b-84aa-5b2f-d532614ed347%28Office.15%29.aspx)|

## Properties

|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b3c44d53-82d5-dcfd-a9f7-c2aee2aa7358%28Office.15%29.aspx)|
|[AutoScaling](http://msdn.microsoft.com/library/fecafb42-56fb-3c33-dc03-cb290b4a28df%28Office.15%29.aspx)|
|[BackWall](http://msdn.microsoft.com/library/c72de543-7be9-55ff-20d0-a5330ca92144%28Office.15%29.aspx)|
|[BarShape](http://msdn.microsoft.com/library/46ce2a4f-8465-493b-ff89-9ddc5e619bf4%28Office.15%29.aspx)|
|[CategoryLabelLevel](http://msdn.microsoft.com/library/b3a54685-18d7-8c24-b2e8-f3bfb03fc69e%28Office.15%29.aspx)|
|[ChartArea](http://msdn.microsoft.com/library/125d6176-b770-900b-8572-ce33b95ad897%28Office.15%29.aspx)|
|[ChartColor](http://msdn.microsoft.com/library/a2bd828b-cf03-2927-8fe6-70414dafd46a%28Office.15%29.aspx)|
|[ChartStyle](http://msdn.microsoft.com/library/b4bc3251-6afc-18e4-214a-a755a46776ba%28Office.15%29.aspx)|
|[ChartTitle](http://msdn.microsoft.com/library/3a083c1f-7a3f-3368-c547-297f0e5d26cb%28Office.15%29.aspx)|
|[ChartType](http://msdn.microsoft.com/library/532a2988-babf-b51a-7548-2f11f94c82a6%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/88872dad-53b2-580a-9bbc-6a29066352a6%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/2e80075a-d113-a602-d09f-c04f6e0d568d%28Office.15%29.aspx)|
|[DataTable](http://msdn.microsoft.com/library/e977daf1-45a1-a069-3d6c-afbe13724d11%28Office.15%29.aspx)|
|[DepthPercent](http://msdn.microsoft.com/library/3b53544f-8800-c1c9-6615-c601d213daee%28Office.15%29.aspx)|
|[DisplayBlanksAs](http://msdn.microsoft.com/library/b4e18939-6214-25e8-a0cd-c984b9f82346%28Office.15%29.aspx)|
|[Elevation](http://msdn.microsoft.com/library/44dde783-5bf7-7c5c-475b-0666337249d7%28Office.15%29.aspx)|
|[Floor](http://msdn.microsoft.com/library/7771ab49-b254-f0f0-a21b-596f541ab6c1%28Office.15%29.aspx)|
|[GapDepth](http://msdn.microsoft.com/library/6020490a-1343-5b79-ff7d-197f78061420%28Office.15%29.aspx)|
|[HasAxis](http://msdn.microsoft.com/library/f2df9f16-980d-fd02-3e09-6d6903dbb6c6%28Office.15%29.aspx)|
|[HasDataTable](http://msdn.microsoft.com/library/c29e7606-086e-8549-2259-332d30c1846a%28Office.15%29.aspx)|
|[HasLegend](http://msdn.microsoft.com/library/e791cc18-03a3-1e60-f064-256cdbd6bd2e%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/9aa0e37a-4d1d-1fc3-d5cb-b8869251ff16%28Office.15%29.aspx)|
|[HeightPercent](http://msdn.microsoft.com/library/a95f2b76-57a1-4c04-9f5f-ccd7852d4ab6%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/4f518463-8bb2-caa6-5383-b54d12f20d07%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/2b1166c0-b2e8-e00b-dcc9-9e89b536e241%28Office.15%29.aspx)|
|[Legend](http://msdn.microsoft.com/library/6396ca0f-63b5-3d4a-4f6b-b4e80a1911b3%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/b64d9f0e-6c1d-9d42-5d0e-8c408c057efc%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/3ff78172-884f-4196-f938-75fa12076ccc%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/a0e53eba-c9e9-7997-4765-90debeb8ae5d%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/9a47bfd6-10b5-5f8e-86c2-e56c468de9d8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/2c0db6d3-995a-cc3c-812b-a80761ac76e4%28Office.15%29.aspx)|
|[Perspective](http://msdn.microsoft.com/library/39367c4a-95a7-afe7-b3e4-29e10a88fbd3%28Office.15%29.aspx)|
|[PivotLayout](http://msdn.microsoft.com/library/b621dc49-5321-5426-35cc-386cac251920%28Office.15%29.aspx)|
|[PlotArea](http://msdn.microsoft.com/library/f3c93a06-b398-a60a-d69d-8249652501eb%28Office.15%29.aspx)|
|[PlotBy](http://msdn.microsoft.com/library/69ff0fbe-7954-6808-68fa-cc92b2851dd8%28Office.15%29.aspx)|
|[PlotVisibleOnly](http://msdn.microsoft.com/library/e09aee43-c3f7-9269-f01a-d6298ab780fa%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/c0cf65c3-6e9f-7e04-9161-13ba118f23f1%28Office.15%29.aspx)|
|[PrintedCommentPages](http://msdn.microsoft.com/library/8f98f7af-4e2f-8743-b82b-c84ae83f6fdf%28Office.15%29.aspx)|
|[ProtectContents](http://msdn.microsoft.com/library/03a731a4-9848-dab1-1b49-b3b631c93a77%28Office.15%29.aspx)|
|[ProtectData](http://msdn.microsoft.com/library/29eb3e29-6005-70bd-cb38-053a5d54ed96%28Office.15%29.aspx)|
|[ProtectDrawingObjects](http://msdn.microsoft.com/library/6e65e306-ef55-7e05-41e2-14a1bbc1456e%28Office.15%29.aspx)|
|[ProtectFormatting](http://msdn.microsoft.com/library/71630b7f-6c89-869d-cd5b-d0a7bacd904a%28Office.15%29.aspx)|
|[ProtectionMode](http://msdn.microsoft.com/library/5a9afe8c-df46-cbfe-d692-d4be8f2e505b%28Office.15%29.aspx)|
|[ProtectSelection](http://msdn.microsoft.com/library/a1b9cf7e-8cc3-f9fe-dfcf-c66469741edb%28Office.15%29.aspx)|
|[RightAngleAxes](http://msdn.microsoft.com/library/632aa454-4113-97d3-a80c-eb745a950c6f%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/bf271f86-18c9-ac74-12ab-f90f4353f71d%28Office.15%29.aspx)|
|[SeriesNameLevel](http://msdn.microsoft.com/library/17ada484-943e-502f-a499-077d1e53e6c1%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/73f72671-ac6a-bc11-44cc-a748171d7777%28Office.15%29.aspx)|
|[ShowAllFieldButtons](http://msdn.microsoft.com/library/b5a9dc1a-2c85-eece-b678-2d3509780a46%28Office.15%29.aspx)|
|[ShowAxisFieldButtons](http://msdn.microsoft.com/library/05eff4ce-c06b-b866-b0d7-8733cb51605a%28Office.15%29.aspx)|
|[ShowDataLabelsOverMaximum](http://msdn.microsoft.com/library/1638b7f6-23e5-2fc1-e81b-5b8f54023967%28Office.15%29.aspx)|
|[ShowExpandCollapseEntireFieldButtons](http://msdn.microsoft.com/library/8fc5a821-ab24-2e48-1100-cec590786cd1%28Office.15%29.aspx)|
|[ShowLegendFieldButtons](http://msdn.microsoft.com/library/44f1554c-145b-8600-07c4-40b6891dab2d%28Office.15%29.aspx)|
|[ShowReportFilterFieldButtons](http://msdn.microsoft.com/library/6b7aa6e2-2216-caef-5936-d9c9681b60db%28Office.15%29.aspx)|
|[ShowValueFieldButtons](http://msdn.microsoft.com/library/7997b313-ce87-95eb-3d1e-b9b7b6eda84b%28Office.15%29.aspx)|
|[SideWall](http://msdn.microsoft.com/library/79a6e074-acd1-c14a-02cc-21e549ebffd8%28Office.15%29.aspx)|
|[Tab](http://msdn.microsoft.com/library/bda235b7-d7c1-e901-718e-4d8215433021%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/ce94f2d8-6a02-d857-bd7a-2488c7f6513a%28Office.15%29.aspx)|
|[Walls](http://msdn.microsoft.com/library/fbee1165-7602-4d77-e5b6-8a127783c96e%28Office.15%29.aspx)|

## See also

#### Other resources

[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
