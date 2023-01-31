# Get-ActualCVEsByProduct 
Enumerates all CVEs for a specified product for a specified year-month combination.
Needed PoSh Module MsrcSecurityUpdates will be downloaded and imported.


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

1. +`Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*" -OutputStyle HTML`
3. `Get-ActualCVEsByProduct -ProductTitle "Microsoft Edge*" -OutputStyle GridView`
4. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle GridView`

Output sample 1.:

!+[HTML](https://github.com/BetaHydri/MSRCGetPatches/blob/master/HTML.jpg "HTML Output")<br />

![GridView](https://github.com/BetaHydri/MSRCGetPatches/blob/master/GridView.jpg "GridView Output")
