
$Windows10VertionTable = @{
	"10240" = "Windows 10 for"
	"1607" = "Windows 10 Version 1607"
	"1803" = "Windows 10 Version 1803"
	"1809" = "Windows 10 Version 1809"
	"1903" = "Windows 10 Version 1903"
	"1909" = "Windows 10 Version 1909"
	"2004" = "Windows 10 Version 2004"
	"2009" = "Windows 10 Version 20H2"
}

$WindowsServerVertionTable = @{
#	"xxxx" = "Windows Server 2012"
#	"xxxx" = "Windows Server 2012 R2"
#	"xxxx" = "Windows Server 2016"
	"1809" = "Windows Server 2019"
#	"xxxx" = "Windows Server, version 1903"
#	"xxxx" = "Windows Server, version 2004"
#	"xxxx" = "Windows Server, version 20H2"
}

###############################################
# Windows Update 日を取得する(日本)
###############################################
function GetWindowsUpdateDay([datetime]$TergetDate){

	# 1日の曜日と US Windows Update 日のオフセット ハッシュテーブル
	$DayOfWeek2WUOffset = @{
		[System.DayOfWeek]"Wednesday"	= 13	# 水曜日
		[System.DayOfWeek]"Thursday"	= 12	# 木曜日
		[System.DayOfWeek]"Friday"		= 11	# 金曜日
		[System.DayOfWeek]"Saturday"	= 10	# 土曜日
		[System.DayOfWeek]"Sunday"		= 9 	# 日曜日
		[System.DayOfWeek]"Monday"		= 8 	# 月曜日
		[System.DayOfWeek]"Tuesday" 	= 7 	# 火曜日
	}

	# 年月が指定されていない(default)
	if( $TergetDate -eq $null ){
		# 今の日時
		$TergetDate = Get-Date
	}

	# 1日
	$1stDay = [datetime]$TergetDate.ToString("yyyy/MM/1")

	# US Windows Update 日のオフセット
	$Offset = $DayOfWeek2WUOffset[$1stDay.DayOfWeek]

	if( $Offset -ne $null ){
		# US Windows Update 日
		$WUDayUS = $1stDay.AddDays($Offset)

		# 日本の Windows Update 日(US Windows Update の翌日)
		$WUDay = $WUDayUS.AddDays(1)
	}
	else{
		$WUDay = $null
	}

	return ($WUDay).ToString("yyyy/MM/dd")
}


###############################################
# 指定月に発行されたセキュリティ更新取得
# https://msrc.microsoft.com/update-guide/
###############################################
function GetSecurityUpdateKBs([datetime]$TergetDay){
	# API Key
	# Microsoft Security Update API(@outlook.com、@live.com MSA が必要)
	# https://portal.msrc.microsoft.com/ja-jp/developer
	$APIKey = "7ac1e01ac58840c190f9e519ccf7cb47"

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



###############################################
# Main
###############################################

if( ($PSVersionTable.PSVersion.Major -ge 5) ){
	if(($PSVersionTable.PSVersion.Major -eq 5) -and ($PSVersionTable.PSVersion.Minor -ne 1)){
		echo "このスクリプト実行には Windows PowerShell 5.1 が必要です"
		echo "https://docs.microsoft.com/ja-jp/powershell/scripting/windows-powershell/install/windows-powershell-system-requirements"
		exit
	}
}
else{
	echo "このスクリプト実行には Windows PowerShell 5.1 または PowerShell Core が必要です"
	exit
}



# 今月のWu日
[datetime]$WuDay = GetWindowsUpdateDay

# 今日
$Today = Get-Date

# WU 日前だったら前月
if( $Today -lt $WuDay ){
	$TergetDay = $Today.AddMonths(-1)
}
else{
	$TergetDay = $Today
}

# KB 取得
[array]$Updates = GetSecurityUpdateKBs $TergetDay

$Updates | Sort-Object -Property ProductName, KB -Unique | Export-Csv ~\Documents\KBs.csv -Encoding OEM

# $Updates | Export-Csv C:\Test\WUs.csv -Encoding OEM

# $Updates | Sort-Object -Property KB, ProductID -Unique | Export-Csv C:\Test\WUs.csv -Encoding OEM


# OS のエディション
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$RegKey = "ProductName"
$OS = (Get-ItemProperty $RegPath -name $RegKey -ErrorAction SilentlyContinue).$RegKey

# OS バージョン
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
$RegKey = "ReleaseId"
$OSVertion = (Get-ItemProperty $RegPath -name $RegKey -ErrorAction SilentlyContinue).$RegKey


# Windows 10
if( $OS  -match "Windows 10" ){

	if( $OSVertion -eq $null ){
		$OSVertion = "10240"
	}

	$SelectString = $Windows10VertionTable[$OSVertion]
}
# Windows Server 実装 例(例レベルなので、Windows Server で使うのであればちゃんと実装する必要あり)
elseif( $OS  -match "Windows Server" ){
	if( $WindowsServerVertionTable.ContainsKey($OSVertion) ){
		$SelectString = $WindowsServerVertionTable[$OSVertion]
	}
	else{
		echo "対応していない Windows Server OS バージョンです"
		echo "必要に応じて拡張実装してください"
		exit
	}
}
else{
	echo "対応していない Windows OS バージョンです"
	echo "必要に応じて拡張実装してください"
	exit
}

[array]$WuKBs = $Updates | ? ProductName -Match $SelectString | Sort-Object -Property KB -Unique

echo $WuKBs

[array]$HotFixs = Get-Hotfix

$HitFlag = $false

foreach ( $WuKB in $WuKBs ){
	$KbName = "KB" + $WuKB.KB
	if( $HotFixs.HotFixID -contains $KbName ){
		$HitFlag = $true
	}
	else{
		$HitFlag = $false
	}
}

if( $HitFlag ){
	echo "最新の WU が当たっています"
}
else{
	echo "最新の WU が当たっていません"
}

