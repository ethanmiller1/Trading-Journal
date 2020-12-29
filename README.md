# Trading Journal 
**Version 1.0.0** 

<div align="center">
    <a href="#usage"><img src="https://tlc.thinkorswim.com/center/main/navigation/01/icon/img-release-notes" width="200px"></a>
    <a href="#usage"><img src="https://upload.wikimedia.org/wikipedia/commons/8/86/Microsoft_Excel_2013_logo.svg" width="200px"></a>
    <br>
</div>

> An Excel spreadsheet to collect and analyze institutionally traded stock and option data.

This spreadsheet was designed to (1) help track your trades throughout the year, (2) keep track of the market conditions surrounding the successes or failures surrounding your trades, and (3) keep record of the personal trading rules you applied to them to test whether they are effective trading rules or not.

## Build

The easiest way to access this Trading Journal is simply to download the `Trading Journal.xlxm` file from the repository. However, to build the workbook yourself, follow the following steps:

1. Create an Excel Macro-Enabled Worksheet (Trading Journal.xlsm)
1. `Developer` > `Visual Basic` > Right-click `VBAProject (Trading Journal.xlms)` > `Insert` > `Module`
1. Copy and paste the code from `functions.vb` into the created `Module1` module.

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/module1.gif)

Access the Scrum project management on [Azure Boards](https://dev.azure.com/ethanromans58/Trading%20Journal/_boards/board/t/Trading%20Journal%20Team/Stories)

## Enable function intellisense

Take the following steps to get function tooltips to top up as you're writing them.

1. Download and open the [`ExcelDna.IntelliSense64.xll`](https://github.com/Excel-DNA/IntelliSense/releases/download/v1.1.0/ExcelDna.IntelliSense64.xll) add-in from the Excel-DNA IntelliSense [Releases](https://github.com/Excel-DNA/IntelliSense/releases) page (under "Assets"). (Note: use `ExcelDna.IntelliSense.xll` if your Excel version is 32-bit. Check at `File` > `Account` > `About Excel`.)
1. Open Excel and navigate to `Developer` > `Excel Add-ins` > `Browse` and select ExcelDna.IntelliSense64.xll from the Windows Explorer.

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/intellisenseAddIn.gif)

3. Create a new Worksheet with the name "\_IntelliSense_", and fill in function descriptions as instructed on the [Getting Started](https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started) page.

![](https://github.com/king-melchizedek/Trading-Journal/raw/master/images/customIntellisense.gif)

(Note: Excel must be restarted before changes made to function descriptions will take effect. Be sure to check out the [VBASamples](https://github.com/Excel-DNA/IntelliSense/tree/master/VBASamples) page if you get stuck.)