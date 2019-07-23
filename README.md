---
page_type: sample
products:
- office-excel
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 1/10/2018 1:16:05 PM
---
# Excel Add-in that supports data import and data analysis 
Shows how to create worksheets and tables, hide worksheets, color worksheet tabs, freeze table headers, sort tables, use conditional table formatting, create charts, add trendlines to charts, hide gridlines, include hyperlinks from one table to another, and convert the workbook to a byte array. 


## Table of Contents
* [Change History](#change-history)
* [Prerequisites](#prerequisites)
* [Get started](#get-started)
* [Build and Test](#build-and-test)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)


## Change History

* January 10, 2018: Initial version.

## Prerequisites

- Visual Studio 2017
- Node JS 
- npm (gets installed when Node JS is installed)
- A git client (instructions below assume you are using a git CLI, such as git bash)
- Office 2016, Version 1710, build 16.0.8605.1000 Click-to-Run, or later. You many need to be an Office Insider to obtain this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## Get started

1. Install the prerequisites above.
2. Install the NPM Package Manager for Visual Studio from [NPM TaskRunner](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner).
3. Clone this repo.
4. Open a Node-enabled system prompt (or git bash) in the root folder of the **Excel-Add-in-SalesLeadsWeb** project (one level down from the root of the solution).
5. Run `npm install`. This will install babel and WebPack.
6. Open the *.sln file in the root of the project in Visual Studio.
7. In Visual Studio, select **View** | **Other Windows** | **Task Runner Explorer**. 
8. In the **Task Runner Explorer**, open **package.json** and **Custom**. 
9. Right-click **build**, and then in the context menu, select **Bindings** | **After build**.

## Build and Test

1. Press F5 to start the add-in.
2. Open the add-in from the **Sales Leads** button on the **Home** ribbon.
3. On the taskpane, click **Import Data**. Two worksheets are created named **Opportunities** and **Customers**. Note that for both worksheets, the tabs are colored red and the top row remains visible if you scroll down in the worksheet.
4. Select a cell with the name **Sally Jump** on the **Opportunities** sheet. (This is the only salesperson with enough data to make a meaningful analysis.)
5. Press **Analyse**. A **Report** worksheet is created. Note that:
> -  Its gridlines are hidden.
> -  The table at the top contains the sales leads for the salesperson "Sally Jump" taken from the **Opportunities** worksheet.
> -  The column of customers in the table are hyperlinks to the customer's contact info in the **Customers** worksheet. 
> -  The **Expected Value** column is conditionally formatted with a color scale, so that the lead with the highest potential value is red, the lead with the lowest is blue, and other leads other leads are mixtures of red, blue, and yellow depending on how close they are to the highest and lowest potential values.
> - The chart shows the sales in 2016 of the salesperson "Sally Jump". 
> - The chart has a trendline showing the trend of the salesperson's sales in 2016.
6. Press **Get File as Base 64 String**. After a few seconds, a long base 64 string will appear in a textbox on the taskpane. To verify that this is the file, copy and paste the string to a website where you can decode and download it as an *.xslx file, such as https://www.base64decode.org/. 
7. Open the downloaded file. It should be identical to the file on which you are running the add-in.

> **Note:** The **Download Report** button is intended only to illustrate a possible enhancement. It is not implemented.

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). If your question is about the Office JavaScript APIs, make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright
Copyright (c) 2018 Microsoft Corporation. All rights reserved.
