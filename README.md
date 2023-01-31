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
2. `Get-ActualCVEsByProduct -ProductTitle "Microsoft Edge*" -OutputStyle GridView`
3. `Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle GridView`
