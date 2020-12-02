###############################################
# 指定月に発行されたセキュリティ更新取得
# https://msrc.microsoft.com/update-guide/
###############################################
function GetSecurityUpdateKBs([string]$APIKey, [datetime]$TergetDay){

	# API Key
	# Microsoft Security Update API(@outlook.com、@live.com MSA が必要)
	# https://portal.msrc.microsoft.com/ja-jp/developer

	# セキュリティ更新を格納するオブジェクト
	class UpdateClass {
		[string] $CVE			# CVE 番号
		[string] $ReleaseDate	# リリース日
		[string] $KB			# KB 番号
		[string] $SubType		# 更新の種類
		[string] $ProductID		# プロダクト ID
		[string] $ProductName	# プロダクト名
		[string] $URI			# KB の URI
	}

	# 月の略称
	$MonthTable = @{
		1 = "Jan"
		2 = "Feb"
		3 = "Mar"
		4 = "Apr"
		5 = "May"
		6 = "Jun"
		7 = "Jul"
		8 = "Aug"
		9 = "Sep"
		10 = "Oct"
		11 = "Nov"
		12 = "Dec"
	}

	if( ($APIKey -eq [string]$null) -or ($TergetDay -eq $null) ){
		return $null
	}

	$TergetMonth = [string]$TergetDay.Year + "-" + $MonthTable[$TergetDay.Month]

	try{
		Import-Module MSRCSecurityUpdates -ErrorAction Stop
	}
	catch{
		echo "必要モジュールがインストールされていません"
		echo "以下手順でモジュールをインストールしてください(要管理権限)"
		echo "Install-Module MSRCSecurityUpdates -Force"
		return $null
	}

	Set-MSRCApiKey -ApiKey $APIKey
	$cvrfDoc = Get-MsrcCvrfDocument -ID $TergetMonth

	# プロダクト名
	$FullProductName = $cvrfDoc.ProductTree.FullProductName


	# プロダクト別のセキュリティ更新
	$Updates = @()

	# セキュリティ更新
	[array]$Vulnerabilitys = $cvrfDoc.Vulnerability

	# CVE ごとの処理
	foreach( $Vulnerability in $Vulnerabilitys){

		$CVE = $Vulnerability.CVE

		# リリース日
		[datetime]$ReleaseDate = ($Vulnerability.RevisionHistory | Sort-Object -Property Date)[0].Date

		# セキュリティ更新の抽出
		[array]$Remediations = $Vulnerability.Remediations | ? Type -eq 2

		# セキュリティ更新ごとの処理
		foreach( $Remediation in $Remediations ){

			# KB 番号が数値のデータの未処理
			if($Remediation.Description.Value -match "[0-9]+"){

				# KB 番号
				$KB = $Remediation.Description.Value

				# 更新の種類
				$SubType = $Remediation.SubType

				# プロダクトの展開
				[array]$ProductIDs = $Remediation.ProductID

				# プロダクト後の処理
				foreach( $ProductID in $ProductIDs ){
					# 各値のセット
					$Update = New-Object UpdateClass

					$Update.CVE = $CVE
					$Update.ReleaseDate = $ReleaseDate.ToString()
					$Update.KB = $KB
					$Update.SubType = $SubType
					$Update.ProductID = $ProductID
					$Update.ProductName = ($FullProductName | ? ProductID -eq $ProductID).Value
					$Update.URI = "https://support.microsoft.com/ja-jp/help/" + $KB

					$Updates += $Update
				}
			}
		}
	}

	return $Updates
}
