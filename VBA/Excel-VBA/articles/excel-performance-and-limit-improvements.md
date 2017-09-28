---
title: Excel performance: Performance and limit improvements
description: This article discusses performance improvements in Microsoft Excel 2016 and Microsoft Excel 2010. This article is one of three companion articles about techniques that you can use to improve performance in Excel as you design and create worksheets.
ms.date: 9/28/2017
---

# Excel performance: Performance and limit improvements

## Excel 2016 performance improvements
<a name="xl2016PerfImp"> </a>

Microsoft Office has fundamentally changed its development and release methodology with Office 2016 and Office 365. Updates and new features are released on a regular cycle, so it becomes important to note the build number that improvements were released, and to check which build number you are using. The timescale that updates become available to you depends on which update option you are using:
- Insider
- Monthly Channel
- Semi-annual Channel

For more information about the Office 2016 release cadence names, see [Slow - Fast Level Names](https://support.office.com/en-US/article/Changes-to-Slow-Fast-level-names-for-Office-Insider-for-Windows-desktop-055ee4f9-9ce3-4fb8-8a9a-ca6745867d52).

The following sections discuss some of the features that have been introduced in Excel 2016 that you can use to improve performance with large or complex workbooks.

### Large Address Aware (LAA) Memory improvement for 32-bit Excel
Although 64-bit Excel has extremely large virtual memory limits, 32-bit Excel has been limited to 2 gigabytes (GB). Many Excel customers have found it difficult to migrate to 64-bit Excel because they use third-party add-ins and controls that are not available in 64-bit versions.

LAA has now been enabled for 32-bit versions of Excel 2013 and Excel 2016, and will minimize out-of-memory error messages. LAA doubles available virtual memory from 2 GB to 4 GB when using 64-bit Windows, and increases available virtual memory from 2 GB to 3 GB under 32-bit Windows.

For more information, see [LAA Capability Change for Excel](https://support.microsoft.com/en-ca/help/3160741/large-address-aware-capability-change-for-excel "LAA Capability Change for Excel"). 

To download a tool that shows how much virtual memory is available and how much is being used, see [Excel Memory Checking Tool](https://fastexcel.wordpress.com/2016/11/27/excel-memory-checking-tool-using-laa-to-increase-useable-excel-memory/).

### Full column references
Previously, workbooks using large numbers of full column references and multiple worksheets (for example `=COUNTIF(Sheet2!A:A,Sheet3!A1)`), might use large amounts of CPU and memory when opened, or rows were deleted. An improvement in Excel 2016 Build 16.0.8212.1000 substantially reduces the memory and CPU used in these circumstances.
 
*My test on a workbook with 6 million formulas using full column references failed with an Out of Memory message at 4 GB of virtual memory with Excel 2013 LAA, but only used 2 GB virtual memory with Excel 2016*.

### Structured references
In some circumstances, editing Excel tables where formulas in the workbook use structured references to the table could be slow with Excel 2013 and previous versions. This led to the perception that tables should not be used with large numbers of rows. Excel 2016 has now fixed this problem. 

*My test showed an editing operation that took 1.9 seconds in Excel 2013 took about 2 milliseconds in Excel 2016.*

For more information, see [Why Structured References are Slow in Excel 2013 but fast in Excel 2016](https://fastexcel.wordpress.com/2017/02/19/why-structured-references-are-slow-in-excel-2013-but-fast-in-excel-2016/).

### Filtering, sorting, copy/pasting
The Excel 2016 team studied a number of large workbooks that show slow response when using filtering, sorting, and copy/pasting, and a number of improvements have been made.

In Excel 2013, after filtering, sorting, or copy/pasting many rows, Excel could be slow responding or would hang. Performance was highly dependent on the count of all rows between the top visible row and the bottom visible row. An improvement made to the internal calculation of vertical user interface positions in Build 16.0.8431.2058 has made these operations much faster. 

Opening a workbook with many filtered or hidden rows, merged cells or outlines could cause high CPU load. A fix in this area was introduced in Build 16.0.8229.1000.

In the past, you could see very slow response after pasting a copied column of cells from a table with filtered rows where the filter resulted in a large number of separate blocks of rows. This area has been substantially improved in Build 16.0.8327.1000.

*My test on copy/pasting 22,000 rows filtered from 44,000 rows showed a dramatic improvement:*
- *For a table, the time went from 39 seconds in Excel 2013 to 2 seconds in Excel 2016.*
- *For a range, the time went from 30 seconds in Excel 2013 to virtually instantaneous in Excel 2016.*

### Copying conditional formats
In Excel 2013, copy/pasting cells containing conditional formats could be slow. This has been significantly improved in Excel 2016 Build 16.0.8229.0.

*My test on copying 44,000 cells with a total of 386,000 conditional format rules showed a substantial improvement:*
- *Excel 2013: 68 seconds*
- *Excel 2016: 7 seconds*

### Adding and deleting worksheets
My test on Excel 2016 Build 16.0.8431.2058 shows a 15-20% speed improvement compared to Excel 2013 when adding and deleting large numbers of worksheets.

### New functions

Excel 2016 Build 16.0.7920.1000 introduced several new and very useful worksheet functions:

- **MAXIFS** and **MINIFS** extend the **COUNTIFS/SUMIFS** family of functions. These functions have good performance characteristics and should be used to replace equivalent array formulas.
- **SWITCH** and **IFS** provide ways of simplifying complex IF statements.
- **TEXTJOIN** and **CONCAT**  let you easily combine text strings from ranges of cells.

### Integrated PowerPivot and Power Query (Get & Transform)
In Excel 2016, PowerPivot and Power Query (Get and Transform) are fully integrated to Excel rather than being add-ins, resulting in improved performance and control.

### Other updates to Excel 2016 for Windows
You can find more details of all the other month-by-month improvements that have been made to Excel 2016 at
[What's new in Excel 2016 for Windows](https://support.office.com/en-gb/article/What-s-new-in-Excel-2016-for-Windows-5fdb9208-ff33-45b6-9e08-1f5cdb3a6c73).


## Excel 2010 performance improvements
<a name="xl2010PerfImp"> </a>

The following sections discuss some features that were introduced in Excel 2010 that you can use to improve performance.

### Feature improvements

Based on user feedback about Excel 2007, Excel 2010 introduced improvements to several features.

|**Feature**|**Improvement**|
|:-----|:-----|
|**Printer and Page Layout View** <br/> |To improve performance of basic user interactions in page layout view, such as entering data, working with formulas or setting margins, Excel 2010 caches the printer settings and introduces optimized rendering calculations. Caching the printer settings reduces the number of network calls and reduces the dependency on a slow or unresponsive printer. In addition, connecting to the printer is cancelable so that the user does not have to wait for a slow or unresponsive printer.  <br/> |
|**Charts** <br/> |Starting in Excel 2010, the rendering speed of charts has increased, especially with large data sets, and text-rendering performance has improved. In addition, Excel 2010 caches an image of a chart and uses the cached version when possible, to avoid unnecessary calculations and rendering.  <br/> |
|**VBA Solutions** <br/> |Improvements to the object model and the way it interacts with Excel increases the performance speed of many VBA solutions when run in Excel 2010 compared with Excel 2007.  <br/> |

### Large data sets and 64-bit Excel

The 64-bit version of Excel 2010 is not constrained to 2 GB of RAM like 32-bit applications. Therefore, 64-bit Excel 2010 enables users to create much larger workbooks. 64-bit Windows enables a larger addressable memory capacity, and 64-bit Excel is designed to take advantage of that capacity. For example, users are able to fill more of the grid with data than was possible in previous versions of Excel. As more RAM is added to the computer, Excel uses that additional memory, allows larger and larger workbooks, and scales with the amount of RAM available.

In addition, because 64-bit Excel enables larger data sets, both 32-bit and 64-bit Excel 2010 introduce improvements to common large data set tasks such as entering and filling down data, sorting, filtering, and copying and pasting data. Memory usage is also optimized to be more efficient, in both the 32-bit and 64-bit versions of Excel. 
 
For more information about the "Big Grid," see  [The "Big Grid" and Increased Limits in Excel 2007](#Office2007excelPerf_BigGridIncreasedLimitsExcel). For more information about the 64-bit version of Office 2010, see [Compatibility Between the 32-bit and 64-bit Versions of Office 2010](http://msdn.microsoft.com/library/24acd0f0-1d3a-435e-8b76-44820648ab54%28Office.14%29.aspx).

### Shapes
<a name="Shapes"> </a>

Excel 2010 introduces significant improvements in the performance of graphics in Excel. At a high level, these improvements are in two areas: scalability and rendering. 

The scalability improvements have a large impact in Excel scenarios because of the large number of graphics contained on worksheets. Often, this large number of shapes is created accidentally by copying and pasting data from a website, or by commonly run automation that creates shapes, but never removes them. This large number of graphics, combined with the way that graphics relate to the data grid in Excel, presents several unique performance challenges. Improvements in Excel 2010 increase the performance speed for worksheets that contain many shapes. 

In addition, starting in Excel 2010, support for hardware acceleration improves rendering. Excel 2010 also introduces performance improvements to the **Select** method of the **Shape** object in the VBA object model.

|**Feature**|**Improvement**|
|:-----|:-----|
|**Basic Use** <br/> |The first set of improvements made in Excel 2010 surrounds basic use scenarios. These scenarios include operations and features such as sorting, filtering, inserting or resizing rows or columns, or merging cells. When these operations occur, it may be necessary to update the position of a graphic object on the grid. In the worst-case scenario, it is necessary to make an update to every single object on the worksheet. In Excel 2010, performance of these basic scenarios improves even when there are thousands of objects on the worksheet. These improvements were not achieved with a single feature or fix, but through a dedicated focus on performance that included improving the shape lookup mechanism, testing stress files, and investigating obstructions.  <br/> |
|**Text Links** <br/> |A text link on a shape is created when the user specifies a formula, for example "=A1", that defines the text for a given shape. These particular shapes were prone to cause performance issues on sheets with a large number of objects and/or when changes were made to cell content. Starting in Excel 2010, the way Excel tracks and updates these shapes has improved to optimize performance for changing cell content. This work improves scenarios such as typing a new value in a cell or performing complex object model operations.  <br/> |
|**Big Grid** <br/> |Starting in Excel 2007, the size of the grid expanded from 65,000 rows to over one million rows. This increase caused some performance and rendering issues when working with graphics objects in the new regions of the larger grid. Starting in Excel 2010, Excel optimizes functionality that relies on using the top left of the grid as the origin to improve the experience of working with graphics in the new regions of the grid. Rendering fidelity and performance are improved relative to Excel 2007.  <br/> |
|**Rendering: Hardware Acceleration** <br/> |Starting in Excel 2010, improvements were made to the graphics platform by adding support for hardware acceleration when rendering 3-D objects. While the GPU can render these objects faster than the CPU, the experience in Excel 2010 depends on the content on your worksheet. If you have a sheet full of 3-D shapes, you will see more benefit from the hardware acceleration improvements than on a worksheet with only 2-D shapes (which do not leverage the GPU).  <br/> |

### Calculation improvements
<a name="Shapes"> </a>

Starting in Excel 2007, multithreaded calculation improved calculation performance. For more information, see [Multithreaded Calculation](#MultithreadedCalculation). Starting in Excel 2010, additional performance improvements were made to further increase calculation speed. Excel 2010 can call user-defined functions asynchronously. Calling functions asynchronously improves performance by allowing several calculations to run at the same time. When you run user-defined functions on a compute cluster, calling functions asynchronously enables several computers to be used to complete the calculations. For more information about asynchronous user-defined functions, see  [Asynchronous User-Defined Functions](http://msdn.microsoft.com/library/142eb27e-fb6f-4da3-bfb7-a88115bbb5d5%28Office.14%29.aspx).

### Multi-core processing
<a name="Shapes"> </a>

Additional investments were made to take advantage of multi-core processors and increase performance for routine tasks. Starting in Excel 2010, the following features use multi-core processors: saving a file, opening a file, refreshing a PivotTable (for external data sources, except OLAP and SharePoint), sorting a cell table, sorting a PivotTable, and auto-sizing a column.

For operations that involve reading and loading or writing data, such as opening a file, saving a file, or refreshing data, splitting the operation into two processes increases performance speed. The first process gets the data, and the second process loads the data into the appropriate structure in memory or writes the data to a file. In this way, as soon as the first process begins reading a portion of data, the second process can immediately start loading or writing that data, while the first process continues to read the next portion of data. Previously, the first process had to finish reading all the data in a certain section before the second process could load that section of the data into memory or write the data to a file.

### PowerPivot
<a name="Shapes"> </a>

PowerPivot refers to a collection of applications and services that provide an end-to-end approach for creating data-driven, user-managed business intelligence solutions in Excel workbooks. PowerPivot for Excel is a data analysis tool that delivers unmatched computational power directly within Excel. Leveraging familiar Excel features, users can transform large quantities of data from almost any source with amazing speed into meaningful information to get the answers they need in seconds.

PowerPivot also integrates with SharePoint. In a SharePoint farm, PowerPivot for SharePoint is the set of server-side applications, services, and features that support team collaboration on business intelligence data. SharePoint provides the platform for collaborating and sharing business intelligence across the team and larger organization. Workbook authors and owners publish and manage the business intelligence that they develop to their SharePoint sites.

For more information about PowerPivot, see  [PowerPivot Overview](http://msdn.microsoft.com/library/c4c393d3-4856-47ac-ab5f-15da2f240d1d.aspx).

### HPC Services for Excel 2010
<a name="Shapes"> </a>

With a wealth of statistical analysis functions, support for constructing complex analyses, and broad extensibility, Excel 2010 is the tool of choice for analyzing business data. As models grow larger and workbooks become more complex, the value of the information generated increases. However, more complex workbooks also require more time to calculate. For complex analyses, it is common for users to spend hours, days, or even weeks completing such complex workbooks.

One solution is to use Windows HPC Server 2008 to scale out Excel calculations across multiple nodes in a Windows high-performance computing (HPC) cluster in parallel. There are three methods for running Excel 2010 calculations in a Windows HPC Server 2008-based cluster: running Excel workbooks in a cluster, running Excel user-defined functions (UDFs) in a cluster, and using Excel as a cluster service-oriented architecture (SOA) client. For more information about HPC Services for Excel 2010, see [Accelerating Excel 2010 with Windows HPC Server 2008](http://www.microsoft.com/downloads/details.aspx?displaylang=en&amp;FamilyID=a48ac6fe-7ea0-4314-97c7-d6875bc895c5)

## Conclusion
<a name="office2016excelperf_Conclusion"> </a>

Excel 2016 introduces performance and limitation improvements focussed on increasing Excel's ability to efficiently handle large and complex workbooks. All of these improvements allow Excel to scale along with hardware, improving performance as the CPU and RAM capacity of computers expand.
  

## Additional resources
<a name="office2007excelperf_AdditionalResources"> </a>

-  [Excel performance: Improving calculation performance](excel-improving-calcuation-performance.md)
    
-  [Excel performance: Tips for optimizing performance obstructions](excel-tips-for-optimizing-performance-obstructions.md)
    
-  [Excel Developer Portal](http://msdn.microsoft.com/en-us/office/aa905411.aspx)


    
  
