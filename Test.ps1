class UpdateClass {
	[string] $CVE
	[string] $Date
	[string] $KB
	# [string] $Type
	[string] $SubType
	[string] $ProductName
	[string] $URI
}


if( (Get-Module MSRCSecurityUpdates) -eq $null ){
	Install-Module MSRCSecurityUpdates -Force
}

Import-Module MSRCSecurityUpdates

Set-MSRCApiKey -Verbose -ApiKey "7ac1e01ac58840c190f9e519ccf7cb47"
$cvrfDoc = Get-MsrcCvrfDocument -ID 2020-Nov

# プロダクト名
$ProductTree = $cvrfDoc.ProductTree
$FullProductName = $ProductTree.FullProductName
# ProductID でマッチングし、Valueでプロダクトを埋める


$Updates = @()

# セキュリティ更新
[array]$Vulnerabilitys = $cvrfDoc.Vulnerability
foreach( $Vulnerability in $Vulnerabilitys){

	$CVE = $Vulnerability.CVE
	$Date = $Vulnerability.RevisionHistory.Date

	$Remediations = $Vulnerability.Remediations | ? Type -eq 2

	foreach( $Remediation in $Remediations ){
		if($Remediation.Description.Value -match "[0-9]+"){
			$KB = $Remediation.Description.Value
			$Type = $Remediation.Type
			$SubType = $Remediation.SubType

			[array]$ProductIDs = $Remediation.ProductID
			foreach( $ProductID in $ProductIDs ){
				$ProductName = ($FullProductName | ? ProductID -eq $ProductID).Value
				if( $ProductName -match "^Windows" ){
					$Update = New-Object UpdateClass

					$Update.CVE = $CVE
					$Update.Date = $Date.ToString()
					$Update.KB = $KB
					# $Update.Type = $Type
					$Update.SubType = $SubType
					$Update.ProductName = ($FullProductName | ? ProductID -eq $ProductID).Value
					$Update.URI = "https://support.microsoft.com/ja-jp/help/" + $KB

					$Updates += $Update
				}
			}
		}
	}
}

$Updates | Sort-Object -Property KB, ProductName -Unique | Export-Csv C:\Test\WUs.csv -Encoding OEM


<#

WU 日過ぎていたら今月 else 前月

マッチングさせる ProductName をハッシュテーブルで定義し、OS 情報から ProductName を特定する
対象の ProductName のみを抽出し、KB でユニークにすると最新の WU が判明する

$Updates | ? ProductName -Match "Windows 10 Version 1909" | Sort-Object -Property KB -Unique

#>

# $Updates | Export-Csv C:\Test\WUs.csv -Encoding OEM
