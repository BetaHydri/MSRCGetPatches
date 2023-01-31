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

|CVE	|CVE-Title	|ProductName	|Severity	|Impact	|FixedBuild	|KB-ID	|KBType	|KBDownloadUrl3|
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
|CVE-2023-21535	|Windows Secure Socket Tunneling Protocol (SSTP) Remote Code Execution Vulnerability	|Windows Server 2016	|Critical	|Remote Code Execution |	10.0.14393.5648	|5022289	|Security Update	|https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289|
|CVE-2023-21543	|Windows Layer 2 Tunneling Protocol (L2TP) Remote Code Execution Vulnerability	|Windows Server 2016	|Critical	|Remote Code Execution	|10.0.14393.5648	|5022289	|Security Update	|https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289|
|CVE-2023-21546	|Windows Layer 2 Tunneling Protocol (L2TP) Remote Code Execution Vulnerability	|Windows Server 2016	|Critical	|Remote Code Execution	|10.0.14393.5648	|5022289	|Security Update	|https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289 |
|CVE-2023-21548	|Windows Secure Socket Tunneling Protocol (SSTP) Remote Code Execution Vulnerability	|Windows Server 2016	|Critical	|Remote Code Execution |	10.0.14393.5648	|5022289	|Security Update	|https://catalog.update.microsoft.com/v7/site/Search.aspx?q=KB5022289|
