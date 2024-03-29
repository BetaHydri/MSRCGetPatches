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
		Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*"  -Date '2022-Dec' -OutputStyle Console

		.EXAMPLE
		Get-ActualCVEsByProduct -ProductTitle "Windows 10 Version 22H2*"  -Date '2022-Dec' -OutputStyle Console | Format-Table

		.EXAMPLE
		Get-ActualCVEsByProduct -ProductTitle "Windows 10 for x64-based Systems"  -Date '12.2022' -OutputStyle Console | Format-Table

		.EXAMPLE
		Get-ActualCVEsByProduct -ProductTitle "Windows 10 Version 22H2 for x64*"  -Date '12.2022' -OutputStyle Console | Select CVE, CVE-Title, Severity, Impact | Format-Table


	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		$ProductTitle,

		[Parameter(Mandatory = $false, ValueFromPipelineByPropertyName)]
		[DateTime]$Date = (Get-Date -Format yyyy-MMM),

		[Parameter(Mandatory, ValueFromPipelineByPropertyName)]
		[ValidateSet("HTML", "GridView", "Console", "Excel")]
		$OutputStyle
	)
	Begin {
		try {
			### Install the module from the PowerShell Gallery
			if (-not (Get-Module MsrcSecurityUpdates -ListAvailable)) {
				Install-Module MsrcSecurityUpdates -Repository PSGallery -SkipPublisherCheck -Scope CurrentUser -Force
			}
			### Load the module if PowerShell is at least version 5.1
			if ($PSVersionTable.PSVersion -gt [version]'5.1') {
				Import-Module -Name MsrcSecurityUpdates
			}
			else {
				Write-Error "Your PoSH Version is $PSVersionTable.PSVersion you need at least version 5.1"
				Exit
			}
			if ($OutputStyle -eq 'Excel') {
				if (-not (Get-Module ImportExcel -ListAvailable)) {
					Install-Module ImportExcel -Repository PSGallery -SkipPublisherCheck -Scope CurrentUser -Force
				}
				Import-Module -Name ImportExcel
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
					'CVE_Title'     = ((Get-MsrcCveTitle -Vulnerability $ID.Vulnerability -ProductTree $id.ProductTree -ProductFilter $ProductType -CVEFilter $($_.CVE)).CVETitle -split ',')[0]
					'ProductName'   = $_.FullProductName
					'Severity'      = $_.Severity
					'Impact'        = $_.Impact
					'FixedBuild'    = ($_.FixedBuild | ForEach-Object {
							$arr = @(); 
							if ($_ -ne '') { 
								[array]$arr += $_;
								$arr -join ',';
							}
						} | Get-Unique -AsString) -join ',' 
					'KB_ID'         = (@(($_.KBArticle).id | Select-Object -Unique) -join ',')
					'KBType'        = (@(($_.KBArticle).SubType | Select-Object -Unique) -join ",") 
					'KBDownloadUrl' = (@(($_.KBArticle).Url  | Select-Object -Unique) -join ',')
				}
				$ProductNameArray += $ProductObj
			}
			Switch ($OutputStyle) {
			
				'HTML'	{
					if ($ProductNameArray) {
						$Report = $ProductNameArray | Sort-Object -Property Severity, CVE  | ConvertTo-Html
						ConvertTo-Html -Title $Title -Head $Header -Body "$Css $Report" | Out-File $HTMLReport 
						Invoke-Item -Path $HTMLReport
					}
					else {
						Write-Warning "No CVEs on $($Date.ToString("yyyy-MMM")) and for $ProductType were found!"
					}
				}
				'GridView'	{
					if ($ProductNameArray) {
						$ProductNameArray  | Sort-Object -Property Severity, CVE  | Out-GridView -Title "$Title"
					}
					else {
						Write-Warning "No CVEs on $($Date.ToString("yyyy-MMM")) and for $ProductType were found!"
					}
				}
				'Console' {
					if ($ProductNameArray) {
						Write-Host "$Title" -ForegroundColor Green
						$ProductNameArray  | Sort-Object -Property Severity, CVE
					}
					else {
						Write-Warning "No CVEs on $($Date.ToString("yyyy-MMM")) and for $ProductType were found!"
					}
				}
				'Excel' {
					if ($ProductNameArray) {
						
						$excelsrcfile = "$env:USERPROFILE\Desktop\CVEs.xlsx"
						if (Test-Path $excelsrcfile -PathType Leaf) {
							Remove-Item -Path $excelsrcfile -Force
						}
						$excelTableName = "CVEs_$($Date.ToString("MMM_yyyy"))" 
						$data = $ProductNameArray  | Sort-Object -Property Severity, CVE
						$ConditionalText1 = New-ConditionalText -Text 'Critical' -BackgroundColor Red
						$ConditionalText2 = New-ConditionalText -Text 'Important' -BackgroundColor Orange
						$ConditionalText3 = New-ConditionalText -Text 'Moderate' -BackgroundColor Yellow
						$ConditionalText4 = New-ConditionalText -Text 'Low' -BackgroundColor LightBlue
						$ConditionalText = @($ConditionalText1, $ConditionalText2, $ConditionalText3, $ConditionalText4)
						$PivotTableDef = New-PivotTableDefinition -PivotTableName "Pivot_$excelTableName" -PivotRows 'Severity', 'Impact', 'CVE_Title', 'CVE', 'KBType', 'KB_ID' -PivotColumns 'ProductName' -PivotData @{Severity = 'Count' } -PivotTableStyle Medium8 -Activate -ChartType BarClustered -ChartTitle $($excelTableName.Replace('_', ' ')) -ChartHeight 650 -ChartWidth 1200
						$data | Export-Excel -Path $excelsrcfile -WorksheetName $excelTableName -Title $Title -ConditionalText $ConditionalText -PivotTableDefinition $PivotTableDef -NoNumberConversion * -AutoSize -AutoFilter -Show
					}
					else {
						Write-Warning "No CVEs on $($Date.ToString("yyyy-MMM")) and for $ProductType were found!"
					}
				}				
			}
		}
		catch [System.Management.Automation.ParameterBindingException] {
			if ($error[0].Exception.Message -match "The argument ""$($Date.ToString("yyyy-MMM"))"" does not belong to the set") {
				Write-Warning "There are no published CVEs for the selected date: $($Date.ToString("yyyy-MMM"))"
			}
			else { 
				$error[0].Exception.Message
			}
		}
		catch {
			$error[0].Exception.Message
		}
	}
	End {}
}

### Sample calls
#Get-ActualCVEsByProduct -ProductTitle "Windows Server 2016*" -OutputStyle HTML -Date "01.2021"
#Get-ActualCVEsByProduct -ProductTitle "Microsoft SQL Server 2016*" -OutputStyle Excel -Date 06.2022
#Get-ActualCVEsByProduct -ProductTitle "*Sharepoint*" -OutputStyle Console -Date 03.2023
#get-ActualCVEsByProduct -ProductTitle "Windows 8.1*" -OutputStyle GridView -Date 01.2023