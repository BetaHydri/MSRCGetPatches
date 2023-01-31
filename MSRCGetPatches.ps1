
Function Get-MsrcCveTitle {
	<#
		.SYNOPSIS
			Helper function to get CVETitle of a specified CVE number
	 
		.DESCRIPTION
		   Helper function to get CVETitle of a specified CVE number
	 
		.PARAMETER Vulnerability
	 
		.PARAMETER ProductTree
	 
		.EXAMPLE
			(Get-MsrcCveTitle -Vulnerability $ID.Vulnerability -ProductTree $id.ProductTree -ProductFilter $ProductType -CVEFilter $($_.CVE)).CVETitle
	 
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		$Vulnerability,

		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		$ProductTree,
	
		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		$ProductFilter,
	
		[Parameter(Mandatory = $False, ValueFromPipelineByPropertyName)]
		$CVEFilter = "*"
	)
	Begin {}
	Process {
		$Vulnerability | ForEach-Object {

			$v = $_

			$v.ProductStatuses | ForEach-Object {
				$ProductTree.FullProductName  | Where-Object { $_.Value -like $ProductFilter } | Select-Object -ExpandProperty ProductID | ForEach-Object {
					$result = $_
					if (($v | Where-Object { ($v.ProductStatuses.ProductID -contains $result) -AND ($v.cve -eq $CVEFilter) })) {
						[PSCustomObject][Ordered]@{
							"CVETitle" = $v.Title.Value
						}
					}
				}
			}
		}
	}
	End {}
}

Function Get-ActualCVEsByProduct {
	<#
		.SYNOPSIS
			Get CVEs from a specified year-month (e.g.: 2023-Jan) of a given product (e.g.: Windows Server 2022)
	 
		.DESCRIPTION
		   Enumerates all CVEs for a specified product and year-month combination
	 
		.PARAMETER ProductTitle
	 
		.PARAMETER Date

		.PARAMETER OutputStyle
	 
		.EXAMPLE
			Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*" -OutputStyle HTML
	 
		.EXAMPLE
			Get-ActualCVEsByProduct -ProductTitle "Microsoft Edge*" -OutputStyle GridView

		.EXAMPLE
			Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle GridView
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		$ProductTitle,

		[Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
		[DateTime]$Date = (Get-Date -Format yyyy-MMM),

		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		[ValidateSet("HTML", "GridView")]
		$OutputStyle
	)
	Begin {
		try {
			### Install the module from the PowerShell Gallery
			Install-Module -Name MsrcSecurityUpdates -Scope CurrentUser

			### Load the module if PowerShell is at least version 5.1
			if ($PSVersionTable.PSVersion -gt [version]'5.1') {
				Import-Module -Name MsrcSecurityUpdates
			}
		}
		catch {
			Write-Error $error[0].Exception.Message
		}
		$Css = "<style>
		body {
		    font-family: cursive;
		    font-size: 10px;
			color: #000000;
			background: #FEFEFE;
		}
		#title{
			color:#000000;
			font-size: 30px;
			font-weight: bold;
			height: 50px;
		    margin-left: 0px;
		    padding-top: 10px;
		}

		#subtitle{
			font-size: 14px;
			margin-left:0px;
		    padding-bottom: 10px;
		}

		table{
			width:100%;
			border-collapse:collapse;
		}
		table td, table th {
			border:1px solid #000000;
			padding:3px 7px 2px 7px;
		}
		table th {
			text-align:center;
			padding-top:5px;
			padding-bottom:4px;
			background-color:#000000;
		color:#fff;
		}
		table tr.alt td {
			color:#000;
			background-color:#EAF2D3;
		}
		</style>"
	}
	Process {
		try {
			# Environment Variables
			$ID = Get-MsrcCvrfDocument -ID $Date.ToString("yyyy-MMM")
			$ProductType = $ProductTitle

			$HTMLReport = "C:\Temp\HTMLReport" + (Get-Date -Format yyyyMMdd) + ".html"
			$Title = "CVE List for $ProductType on " + ((Get-Culture).DateTimeFormat.GetMonthName(($Date).Month)) + " " + ($Date).year
			$Header = "<div id='title'>$Title</div>$br
			                <div id='subtitle'>Report generated: $(Get-Date)</div>"

			$ProductNameArray = @()
			$ProductObj = @{}
			Get-MsrcCvrfAffectedSoftware -Vulnerability $ID.Vulnerability -ProductTree $ID.ProductTree |  Where-Object { $_.FullProductName -like $ProductType } | ForEach-Object {
			
				$ProductObj = [PSCustomObject][Ordered]@{
					'CVE'           = $_.CVE
					'CVE-Title'     = ((Get-MsrcCveTitle -Vulnerability $ID.Vulnerability -ProductTree $id.ProductTree -ProductFilter $ProductType -CVEFilter $($_.CVE)).CVETitle -split ',')[0]
					'ProductName'   = $_.FullProductName
					'Severity'      = $_.Severity
					'Impact'        = $_.Impact
					'FixedBuild'    = ($_.FixedBuild -split ',')[0]
					'KB-ID'         = ($_.KBArticle).ID
					'KBType'        = ($_.KBArticle).SubType
					'KBDownloadUrl' = ($_.KBArticle).Url
				}
				$ProductNameArray += $ProductObj
			}
			Switch ($OutputStyle) {
			
				'HTML'	{
					$Report = $ProductNameArray | Sort-Object -Property Severity, CVE  | ConvertTo-Html
					ConvertTo-Html -Title $Title -Head $Header -Body "$Css $Report" | Out-File $HTMLReport 
					Invoke-Item -Path $HTMLReport
				}
				'GridView'	{
					$ProductNameArray  | Sort-Object -Property Severity, CVE  | Out-GridView -Title "$Title"
				}
			
			}
		}
		catch {
			$error[0].Exception.Message
		}
	}
	End {}
}

Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016" -OutputStyle HTML
#Get-ActualCVEsByProduct -ProductTitle "*" -OutputStyle HTML