using module ".\GetSecurityUpdate.psm1"


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


$APIKey = "7ac1e01ac58840c190f9e519ccf7cb47"

# KB 取得
[array]$Updates = GetSecurityUpdateKBs $APIKey $TergetDay

$Updates | Sort-Object -Property ProductID, KB -Unique | Export-Csv ~\Documents\KBs.csv -Encoding OEM

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

