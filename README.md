# pptx-importer
A Python program that exports research-selected data into formatted PowerPoint charts and tables.

## Definitions & Concepts

### Static Data Reports
These are reports which always feature the same number of tabs and the same number of columns in every import document. 
The number of rows indicates how many pages, or report files, are generated. 
If the number of tabs or columns differ from the expectations of the software, errors will be thrown.
Additionally, many of these reports rely on name-based indexing, so name-spelling mismatches in the data, may throw errors.

### Flexible Data Reports
These reports generate one chart for every tab with color-coded data selections. 
This approach allows for maximum flexibility and control over the output.
Because the selection is a manual process, errors often occur based on selections.
Common errors include data array mismatches (the data selection is missing or has too many values)
and too much data being selected for a single chart. The software will not assemble mutli-level axes.

### Combinations
A combination of the above approaches is possible, and is used for Country Reports, 
which use the columns in each tab to generate a report file for each country. 
Otherwise, the color selection method is used for selecting data. 
To test for errors, the report is run as a standard general report, then as a country report.

### The necessity of breakpoints
Python is not smart enough on its own to select and separate data, so the software revolves around breakpoints: one tab per chart, color selections, and report-specific, pre-determined columns.
This is a core tenant of the software and moving away from it would be an expensive enterprise.

### PPTX is a markup-based program
PowerPoint files are zip files of XML data. 
The software edits the XML for PowerPoint's software to read and render.
In this way, PPTX files are like HTML files.
Unlike HTML, XML is not forgiving, especially for PPTX files, which rely on a specific order of nested elements.
The Python-PPTX library seeks to abstract this process, but some styles and functions require direct editing of the XML.
If direct editing is needed, create two PowerPoint files, one styled in the desired way, the other not. 
[Unzip the files to extract the XML](https://support.microsoft.com/en-us/office/extract-files-or-objects-from-a-powerpoint-file-85511e6f-9e76-41ad-8424-eab8a5bbc517/), then compare the two versions of the file to see how the XML should be structured.
Please note: pages and charts are separate XML files. 
Also, it's easiest to compare one thing at a time, as the files are labelled by index, making finding the preferred file a challenge.

### Room for growth
At this stage, the file structure is more divided and robust than is needed for the project.
The logic of dividing the functions into specific pigeonholes is to allow the software to grow into new territory, without requiring major refactors down the line.
The file structure logic is thus:
* **Data Handlers**: This contains functions and classes which allow for data to be collected and contained in structures which allow for report assembly. This could eventually be used for data coming from non-excel locations.
* **PPTX Handlers**: This contains functions and classes which are specifically related to the Python-PPTX library and the manipulation of PPTX files. (e.g., the *text_hanlders.py* file is about how PowerPoint handles and styles text, not the manipulation of text itself.)
* **Utilities**: This folder contains the files necessary for the software to function, including the default template, as well as general functions used in the other packages.
* **XLSX Handlers**: This contains functions and classes specifically related to the manipulation of Excel files. There is a lot of room for growth here.

## Process

### Selections are made in the GUI
*__main\__.py* generates the GUI and allows users to make selections.
It also drives the below process by calling specific functions in order. 
When the process concludes, it saves the resulting PPTX file.

### The PPTX Template is Processed
*pptx_handlers/template_reader.py* reads the template as found in the GUI settings tab, pulling metadata on each layout into a class instance, including:
* The layout name
* The count and type of placeholders
* The presence of special placeholders, like titles, footers, and page numbers

This metadata is used when assembling reports to select the correct layout, based either on name—for static data reports—or on the placeholder count and type—for flexible data reports.

### The Excel File is Processed and the Report Structure Assembled
*data_hanlders/report_outlines.py* is used to create class instances that contain the data for the report. 
Each class type relates to the type of report and its specific structure. 
For general data reports, the structure is simple: a list of page class instances, each with its own list of chart instances.
For static data reports, each page must be created with very specific parameters and data cuts.

#### Report Class
The report classes act as containers which contain the report metadata and the pages to include in the report

#### Page Class
The page class contains page metadata and a list of the charts, shapes, and text to include on that page.
This class is found in a list of Page Class Instances in the Report Class.

#### Chart Class
The chart class contains a single chart's metadata (like chart type and style) and the actual data.
Chart class instances are contained as a list in the Page class

#### Chart Style Class
Unlike the chart metadata pulled from the tab selections, or report-specific parameters, 
chart styles are predefined in *pptx_handlers.chart_definitions.py*.
These class instances are sets of predefined, PPTX-specific style information chosed to match Technomic's preferred styles.
An instance of this class is chosen based on chart type and attached to the Chart Class

### The New PPTX File is Assembled
*pptx_handlers/pptx_creator.py* iterates over a provided Report Class instance, using the Python-PPTX library to create and style pages and charts.
It will look at each page's metadata and number of charts to determine the correct layout,
then it will insert charts by iterating over the list of chart instances in the page instance.
Each chart is styled according to report selections and the predefined chart styles.