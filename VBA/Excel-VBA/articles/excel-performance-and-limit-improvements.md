---
title: Excel performance - Performance and limit improvements
description: Find out about performance improvements in Excel 2016 and Excel 2010. 
ms.date: 10/06/2017 
author: FastExcel
---

# Excel performance: Performance and limit improvements

**Applies to:** Excel | Excel 2016 | Excel 2013 | Excel 2010 | Office 2016 | SharePoint Server 2010 | VBA

Excel 2016 introduces new features that you can use to improve performance when you're working with large or complex Excel workbooks

## LAA memory improvement for 32-bit Excel

Although the 64-bit version of Excel has large virtual memory limits, the 32-bit version has only 2 GBs of virtual memory. Some customers use the 32-bit version because some third-party add-ins and controls are not available in the 64-bit version.

The 32-bit versions of Excel 2013 and Excel 2016 now have Large Address Aware (LAA) enabled. This will minimize out-of-memory error messages.

LAA doubles available virtual memory from 2 GB to 4 GB on 64-bit versions of Windows, and increases available virtual memory from 2 GB to 3 GB on 32-bit versions of Windows.

For more information, see [Large Address Aware Capability Change for Excel](https://support.microsoft.com/en-ca/help/3160741/large-address-aware-capability-change-for-excel).

To download a tool that shows how much virtual memory is available and how much is being used, see [Excel Memory Checking Tool](https://fastexcel.wordpress.com/2016/11/27/excel-memory-checking-tool-using-laa-to-increase-useable-excel-memory/).

## Full column references

In earlier versions of Excel, workbooks using large numbers of full column references and multiple worksheets (for example `=COUNTIF(Sheet2!A:A,Sheet3!A1)`) might use large amounts of memory and CPU when opened or when rows were deleted. 

Excel 2016 Build 16.0.8212.1000 reduces the memory and CPU used in these circumstances.
 
*In a sample test on a workbook with 6 million formulas, using full column references failed with an out-of-memory message at 4 GB of virtual memory with Excel 2013 LAA and with Excel 2010, but only used 2 GB of virtual memory with Excel 2016*.

## Structured references

In Excel 2013 and earlier versions, editing tables where formulas in the workbook use structured references to the table was slow. This led to the perception that tables should not be used with large numbers of rows. This issue no longer occurs in Excel 2016.

*For example, an editing operation that took 1.9 seconds in Excel 2013 and Excel 2010 took about 2 milliseconds in Excel 2016.*

## Filtering, sorting, and copy/pasting

We've made a number of improvements to the response time when filtering, sorting, and copy/pasting in large workbooks.

In Excel 2013, after filtering, sorting, or copy/pasting many rows, Excel could be slow responding or would hang. Performance was dependent on the count of all rows between the top visible row and the bottom visible row. These operations are much faster after we improved the internal calculation of vertical user interface positions in Build 16.0.8431.2058. 

Opening a workbook with many filtered or hidden rows, merged cells, or outlines could cause high CPU load. We introduced a fix in this area in Build 16.0.8229.1000.

After pasting a copied column of cells from a table with filtered rows where the filter resulted in a large number of separate blocks of rows, the response time was very slow. This has been improved in Build 16.0.8327.1000.

*A sample test on copy/pasting 22,000 rows filtered from 44,000 rows showed a dramatic improvement:*
- *For a table, the time went from 39 seconds in Excel 2013 and 18 seconds in Excel 2010 to 2 seconds in Excel 2016.*
- *For a range, the time went from 30 seconds in Excel 2013 and 13 seconds in Excel 2010 to instantaneous in Excel 2016.*

## Copying conditional formats

In Excel 2013, copy/pasting cells containing conditional formats could be slow. This has been significantly improved in Excel 2016 Build 16.0.8229.0.

*A sample test on copying 44,000 cells with a total of 386,000 conditional format rules showed a substantial improvement:*
- *Excel 2010: 70 seconds*
- *Excel 2013: 68 seconds*
- *Excel 2016: 7 seconds*

## Adding and deleting worksheets

When adding and deleting large numbers of worksheets, a sample test on Excel 2016 Build 16.0.8431.2058 shows a 15%–20% improvement in speed compared to Excel 2013, but 5-10% slower than Excel 2010.

## New functions

Excel 2016 Build 16.0.7920.1000 introduces several useful worksheet functions:

- **MAXIFS** and **MINIFS** extend the **COUNTIFS/SUMIFS** family of functions. These functions have good performance characteristics. Use them to replace equivalent array formulas.
- **TEXTJOIN** and **CONCAT**  let you easily combine text strings from ranges of cells. Use them to replace equivalent VBA UDFs.

## Other updates to Excel 2016 for Windows

For more details about the month-by-month improvements to Excel 2016, see [What's new in Excel 2016 for Windows](https://support.office.com/en-gb/article/What-s-new-in-Excel-2016-for-Windows-5fdb9208-ff33-45b6-9e08-1f5cdb3a6c73).


## Excel 2010 performance improvements

Based on user feedback about Excel 2007, Excel 2010 introduces improvements to several features.

|**Feature**|**Improvement**|
|:-----|:-----|
|**Printer and page layout view** <br/> |To improve performance of basic user interactions in page layout view, such as entering data, working with formulas or setting margins, Excel 2010 caches the printer settings and introduces optimized rendering calculations. Caching the printer settings reduces the number of network calls and reduces the dependency on a slow or unresponsive printer. In addition, connecting to the printer is cancelable so that the user does not have to wait for a slow or unresponsive printer.  <br/> |
|**Charts** <br/> |Starting in Excel 2010, the rendering speed of charts has increased, especially with large data sets, and text-rendering performance has improved. In addition, Excel 2010 caches an image of a chart and uses the cached version when possible, to avoid unnecessary calculations and rendering.  <br/> |
|**VBA solutions** <br/> |Improvements to the object model and the way it interacts with Excel increases the performance speed of many VBA solutions when run in Excel 2010 compared with Excel 2007.  <br/> |

### Large data sets and the 64-bit version of Excel

The 64-bit version of Excel 2010 is not constrained to 2 GB of RAM like the 32-bit version applications nor upto 4 GB of RAM like the Large Address Aware 32-bit version applications. Therefore, the 64-bit version of Excel 2010 enables users to create much larger workbooks. The 64-bit version of Windows enables a larger addressable memory capacity, and Excel is designed to take advantage of that capacity. For example, users are able to fill more of the grid with data than was possible in previous versions of Excel. As more RAM is added to the computer, Excel uses that additional memory, allows larger and larger workbooks, and scales with the amount of RAM available.

In addition, because the 64-bit version of Excel enables larger data sets, both the 32-bit and 64-bit versions of Excel 2010 introduce improvements to common large data set tasks such as entering and filling down data, sorting, filtering, and copying and pasting data. Memory usage is also optimized to be more efficient in both the 32-bit and 64-bit versions of Excel. 

For more information about the 64-bit version of Office 2010, see [Compatibility Between the 32-bit and 64-bit Versions of Office 2010](http://msdn.microsoft.com/library/24acd0f0-1d3a-435e-8b76-44820648ab54%28Office.14%29.aspx) and for choosing between 64-bit and 32-bit, see [Choose between the 64-bit or 32-bit version of Office](https://support.office.com/en-us/article/Choose-between-the-64-bit-or-32-bit-version-of-Office-2dee7807-8f95-4d0c-b5fe-6c6f49b8d261#32or64Bit&32or64Bit).

### Shapes
<a name="Shapes"> </a>

Excel 2010 introduces significant improvements in the performance of graphics in Excel. At a high level, these improvements are in two areas: scalability and rendering. 

The scalability improvements have a large impact in Excel scenarios because of the large number of graphics contained on worksheets. Often, this large number of shapes is created accidentally by copying and pasting data from a website, or by commonly run automation that creates shapes, but never removes them. This large number of graphics, combined with the way that graphics relate to the data grid in Excel, presents several unique performance challenges. Improvements in Excel 2010 increase the performance speed for worksheets that contain many shapes. 

In addition, starting in Excel 2010, support for hardware acceleration improves rendering. Excel 2010 also introduces performance improvements to the **Select** method of the **Shape** object in the VBA object model.

|**Feature**|**Improvement**|
|:-----|:-----|
|**Basic use** <br/> |The first set of improvements made in Excel 2010 surrounds basic use scenarios. These scenarios include operations and features such as sorting, filtering, inserting or resizing rows or columns, or merging cells. When these operations occur, it may be necessary to update the position of a graphic object on the grid. In the worst-case scenario, it is necessary to make an update to every single object on the worksheet. In Excel 2010, performance of these basic scenarios improves even when there are thousands of objects on the worksheet. These improvements were not achieved with a single feature or fix, but through a dedicated focus on performance that included improving the shape lookup mechanism, testing stress files, and investigating obstructions.  <br/> |
|**Text links** <br/> |A text link on a shape is created when the user specifies a formula, for example "=A1", that defines the text for a given shape. These particular shapes were prone to cause performance issues on sheets with a large number of objects and/or when changes were made to cell content. Starting in Excel 2010, the way Excel tracks and updates these shapes has improved to optimize performance for changing cell content. This work improves scenarios such as typing a new value in a cell or performing complex object model operations.  <br/> |
|**Big Grid** <br/> |Starting in Excel 2007, the size of the grid expanded from 65,000 rows to over one million rows. This increase caused some performance and rendering issues when working with graphics objects in the new regions of the larger grid. Starting in Excel 2010, Excel optimizes functionality that relies on using the top left of the grid as the origin to improve the experience of working with graphics in the new regions of the grid. Rendering fidelity and performance are improved relative to Excel 2007.  <br/> |
|**Rendering: Hardware acceleration** <br/> |Starting in Excel 2010, improvements were made to the graphics platform by adding support for hardware acceleration when rendering 3-D objects. While the GPU can render these objects faster than the CPU, the experience in Excel 2010 depends on the content on your worksheet. If you have a sheet full of 3-D shapes, you will see more benefit from the hardware acceleration improvements than on a worksheet with only 2-D shapes (which do not leverage the GPU).  <br/> |
 
### Calculation improvements

Starting in Excel 2007, multithreaded calculation improved calculation performance. 

Starting in Excel 2010, additional performance improvements were made to further increase calculation speed. Excel 2010 can call user-defined functions asynchronously. Calling functions asynchronously improves performance by allowing several calculations to run at the same time. When you run user-defined functions on a compute cluster, calling functions asynchronously enables several computers to be used to complete the calculations. For more information, see [Asynchronous User-Defined Functions](http://msdn.microsoft.com/library/142eb27e-fb6f-4da3-bfb7-a88115bbb5d5%28Office.14%29.aspx).

### Multi-core processing

Excel 2010 made additional investments to take advantage of multi-core processors and increase performance for routine tasks. Starting in Excel 2010, the following features use multi-core processors: saving a file, opening a file, refreshing a PivotTable (for external data sources, except OLAP and SharePoint), sorting a cell table, sorting a PivotTable, and auto-sizing a column.
 
For operations that involve reading and loading or writing data, such as opening a file, saving a file, or refreshing data, splitting the operation into two processes increases performance speed. The first process gets the data, and the second process loads the data into the appropriate structure in memory or writes the data to a file. In this way, as soon as the first process begins reading a portion of data, the second process can immediately start loading or writing that data, while the first process continues to read the next portion of data. Previously, the first process had to finish reading all the data in a certain section before the second process could load that section of the data into memory or write the data to a file.

### PowerPivot

PowerPivot refers to a collection of applications and services that provide an end-to-end approach for creating data-driven, user-managed business intelligence solutions in Excel workbooks. PowerPivot for Excel is a data analysis tool that delivers unmatched computational power directly within Excel. Leveraging familiar Excel features, users can transform large quantities of data from almost any source with amazing speed into meaningful information to get the answers they need in seconds.

PowerPivot also integrates with SharePoint. In a SharePoint farm, PowerPivot for SharePoint is the set of server-side applications, services, and features that support team collaboration on business intelligence data. SharePoint provides the platform for collaborating and sharing business intelligence across the team and larger organization. Workbook authors and owners publish and manage the business intelligence that they develop to their SharePoint sites.
  
For more information about PowerPivot, see  [PowerPivot Overview](http://msdn.microsoft.com/library/c4c393d3-4856-47ac-ab5f-15da2f240d1d.aspx).

### HPC Services for Excel 2010

With a wealth of statistical analysis functions, support for constructing complex analyses, and broad extensibility, Excel 2010 is the tool of choice for analyzing business data. As models grow larger and workbooks become more complex, the value of the information generated increases. However, more complex workbooks also require more time to calculate. For complex analyses, it is common for users to spend hours, days, or even weeks completing such complex workbooks.
 
One solution is to use Windows HPC Server 2008 to scale out Excel calculations across multiple nodes in a Windows high-performance computing (HPC) cluster in parallel. There are three methods for running Excel 2010 calculations in a Windows HPC Server 2008 based cluster: running Excel workbooks in a cluster, running Excel user-defined functions (UDFs) in a cluster, and using Excel as a cluster service-oriented architecture (SOA) client. 

For more information about HPC Services for Excel 2010, see [Accelerating Excel 2010 with Windows HPC Server 2008 R2](http://www.microsoft.com/downloads/details.aspx?displaylang=en&amp;FamilyID=a48ac6fe-7ea0-4314-97c7-d6875bc895c5).

## Conclusion
<a name="office2016excelperf_Conclusion"> </a>

Excel 2016 introduces performance and limitation improvements focused on increasing Excel's ability to efficiently handle large and complex workbooks. These improvements allow Excel to scale along with hardware, improving performance as the CPU and RAM capacity of computers expand.

## Additional resources
<a name="office2007excelperf_AdditionalResources"> </a>

- [Excel performance: Improving calculation performance](excel-improving-calcuation-performance.md)
    
- [Excel performance: Tips for optimizing performance obstructions](excel-tips-for-optimizing-performance-obstructions.md)
    
- [Excel Developer Portal](http://msdn.microsoft.com/en-us/office/aa905411.aspx)
    
- [Changes to Slow/Fast level names for Office Insider for Windows desktop](https://support.office.com/en-US/article/Changes-to-Slow-Fast-level-names-for-Office-Insider-for-Windows-desktop-055ee4f9-9ce3-4fb8-8a9a-ca6745867d52)
    
  
