# Get-ActualCVEsByProduct 
Enumerates all CVEs for a specified product for a specified year-month combination

PARAMETER ***ProductTitle*** (mandatory) <br />
* `"Windows Server 2016*"` <br />
* `"Windows Server 2016"` <br />
* `"Microsoft Edge*"`<br />

PARAMETER ***OutputStyle*** (mandatory)<br />
* `HTML`
* `GridView`

PARAMETER ***Date*** (optional)<br />
* `2021-Jun`
* `2023-Jan`

Samples:

1. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*" -OutputStyle HTML`
3. `Get-ActualCVEsByProduct -ProductTitle "Microsoft Edge*" -OutputStyle GridView`
4. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle GridView`

Output:

|<sub><sup>CVE</sub></sup>	|<sub><sup>CVE-Title</sub></sup>	|<sub><sup>ProductName</sub></sup>	|<sub><sup>Severity</sub></sup>	|<sub><sup>Impact</sub></sup>	|<sub><sup>FixedBuild</sub></sup>	|<sub><sup>KB-ID</sub></sup>	|<sub><sup>KBType</sub></sup>	|<sub><sup>KBDownloadUrl</sub></sup>|
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
|<sub><sup>CVE-2023-21535</sub></sup>	|<sub><sup>Windows Secure Socket Tunneling Protocol (SSTP) Remote Code Execution Vulnerability</sub></sup>	|<sub><sup>Windows Server 2016</sub></sup>	|<sub><sup>Critical</sub></sup>	|<sub><sup>Remote Code Execution</sub></sup> |	<sub><sup>10.0.14393.5648</sub></sup>	|<sub><sup>5022289</sub></sup>	|<sub><sup>Security Update</sub></sup>	|<sub><sup>https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289</sub></sup>|
|<sub><sup>CVE-2023-21543</sub></sup>	|<sub><sup>Windows Layer 2 Tunneling Protocol (L2TP) Remote Code Execution Vulnerability</sub></sup>|<sub><sup>Windows Server 2016</sub></sup>	|<sub><sup>Critical</sub></sup>	|<sub><sup>Remote Code Execution</sub></sup>	|<sub><sup>10.0.14393.5648</sub></sup>	|<sub><sup>5022289</sub></sup>	|<sub><sup>Security Update</sub></sup>	|<sub><sup>https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289</sub></sup>|
|<sub><sup>CVE-2023-21546</sub></sup>	|<sub><sup>Windows Layer 2 Tunneling Protocol (L2TP) Remote Code Execution Vulnerability</sub></sup>	|<sub><sup>Windows Server 2016</sub></sup>	|<sub><sup>Critical</sub></sup>	|<sub><sup>Remote Code Execution</sub></sup>	|<sub><sup>10.0.14393.5648</sub></sup>	|<sub><sup>5022289</sub></sup>	|<sub><sup>Security Update</sub></sup>	|<sub><sup>https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289</sub></sup> |
