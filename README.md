# Get-ActualCVEsByProduct 
Enumerates all CVEs for a specified product for a specified year-month combination.<br />
Needed PoSh module [***msrcSecurityUpdates***](https://github.com/microsoft/MSRC-Microsoft-Security-Updates-API) and [***ImportExcel***](https://github.com/dfinke/ImportExcel) will be downloaded and imported automatically.

After download import module as follows:  
`Import-Module .\MSRCGetPatches\MSRCGetPatches.ps1`

PARAMETER ***ProductTitle*** (mandatory) <br />
* `"Windows Server 2016*"` <br />
* `"Windows Server 2016"` <br />
* `"Microsoft Edge*"`<br />

PARAMETER ***OutputStyle*** (mandatory)<br />
* `HTML`
* `GridView`
* `Excel` by using module ImportExcel from dfinke
* `Console`

PARAMETER ***Date*** (optional)<br />
<sub>If not specified, the current year and month will be used!</sub>
* `2021-Jun`
* `2023-Jan`
* `03/01/2022`
* `11.2022`

Samples:

1. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*" -OutputStyle HTML`
3. `Get-ActualCVEsByProduct -ProductTitle "Microsoft Edge*" -OutputStyle GridView`
4. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle Excel`
5. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle Console`
6. `Get-ActualCVEsByProduct -ProductTitle "Windows 10 for x64-based Systems"  -Date '12.2022' -OutputStyle Console | Format-Table`
7. `Get-ActualCVEsByProduct -ProductTitle "Windows 10 Version 22H2 for x64*"  -Date '12.2022' -OutputStyle Console | Select CVE, CVE-Title, Severity, Impact | Format-Table`
8. `Get-ActualCVEsByProduct -ProductTitle "Microsoft SQL Server 2016*" -OutputStyle GridView -Date 06.2022`

Output sample 1. as HTML Report:

![HTML](https://github.com/BetaHydri/MSRCGetPatches/blob/master/HTML.jpg "HTML Output")<br />

Output sample 1. as GridVIew:

![GridView](https://github.com/BetaHydri/MSRCGetPatches/blob/master/GridView.jpg "GridView Output")<br />

Output sample 3.:

![GridView](https://github.com/BetaHydri/MSRCGetPatches/blob/master/Console.jpg "Console Output")<br />

Output sample 4.:

![Excel](https://github.com/BetaHydri/MSRCGetPatches/blob/master/Excel1.jpg "Excel Table View")

![Excel](https://github.com/BetaHydri/MSRCGetPatches/blob/master/ExcelPivotChart.jpg "Excel Pivot View")
