$ExportStorageSystemCapacity = ''
$ExportVolumeDetail = ''

# ****************************************************************************

$From = "DCS-SRM-Storage Management-CT <postmaster@sanlam.co.za>"
$To = "Troy Hector - BCX <troy.hector@bcx.co.za>"
#$To = "Pierre Esterhuizen (SGT) <Pierre.Esterhuizen@sanlam.co.za>"
#$CC = "DCS-SRM-Storage Management-CT <DCS-SRM-StorageManagement-CT@bcx.co.za>"
$Subject = "SRM Storage Capacity Report - $((Get-Date).ToString('dddd, dd MMM yyyy HH:mm'))"
$SMTPServer = "mail.sanlam.co.za"
$SMTPPort = "25"

$Body = @"
SRM Storage Capacity Report - $((Get-Date).ToString('dddd, dd MMM yyyy HH:mm'))

"@


$Report = "~\SRM_Storage_Capacity_Report_$((Get-Date).ToString('yyyy-MM-dd')).xlsx"

# ****************************************************************************

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

Function Get-StorageSystemCapacity{

    Invoke-WebRequest -Uri https://stor2rrd.mud.internal.co.za/stor2rrd_reports/LATEST/Report_capacity_physical_storage.csv -OutFile .\Report_capacity_physical_storage.csv
    $ExportStorageSystemCapacity = Import-Csv .\Report_capacity_physical_storage.csv -Delimiter ';' |
    Where-Object {(
        $_.'Storage name' -like '*SKY*') -and (
        $_.'Storage Pool name' -notlike 'EXCP00*')
    }|
    Select-Object 'Storage name',
        'Storage Tier',
        'Storage Pool name',
        'Storage Pool total capacity (GB)',
        'Storage Pool allocated capacity (GB)',
        'Storage Pool available capacity (GB)',
        'Backend Storage',
        'Business Unit',
        'Location' | 
    Sort-Object 'Location', 'Storage Tier', 'Storage name', 'Storage Pool name'

    $ExportStorageSystemCapacity | ForEach-Object {
        $_.'Storage Pool total capacity (GB)' = $_.'Storage Pool total capacity (GB)'.replace('.',',')
        $_.'Storage Pool allocated capacity (GB)' = $_.'Storage Pool allocated capacity (GB)'.replace('.',',')
        $_.'Storage Pool available capacity (GB)' = $_.'Storage Pool available capacity (GB)'.replace('.',',')
    } 

    $ExportStorageSystemCapacity | Export-Csv .\Report_capacity_physical_storage_filtered.csv -NoTypeInformation

    $ExcelStorageSystemCapacity = @{
        WorksheetName = 'storage_system_caacity'
        TableName = 'storage_system_caacity'
        TableStyle = 'Medium2'
        AutoSize = $true
        IncludePivotTable = $true
        PivotRows = 'Location','Storage name','Storage Tier', 'Storage Pool name'
        PivotDataToColumn = $true
        PivotData = @{'Storage Pool available capacity (GB)'='Sum';'Storage Pool allocated capacity (GB)'='Sum';'Storage Pool total capacity (GB)'='Sum'}
        # PivotFilter = 'Backend Storage'
    }

    $ExportStorageSystemCapacity | Export-Excel $Report @ExcelStorageSystemCapacity
}
Get-StorageSystemCapacity

Function Get-VolumeDetail{

    Invoke-WebRequest -Uri https://stor2rrd.mud.internal.co.za/stor2rrd_reports/LATEST/Report_volume_storage.csv -OutFile .\Report_volume_storage.csv
    $ExportVolumeDetail = Import-Csv .\Report_volume_storage.csv -Delimiter ';' |
    Where-Object {(
        $_.'Storage System name' -like '*SKY*') -and (
        $_.'Storage System name' -notmatch 'SKY_DX200_S3_BDC|V3700_Sanlam_SKY')
    }|
    Select-Object 'Volume Name',
        'LUN UID',
        'Capacity (GiB)',
        'Storage Pool Name',
        'Storage System name',
        'Storage Tier',
        'Host Mappings' | 
    Sort-Object 'Volume Name'

    $ExportVolumeDetail | ForEach-Object {
        $_.'Capacity (GiB)' = $_.'Capacity (GiB)'.replace('.',',')
    } 

    $ExportVolumeDetail | Export-Csv .\Report_volume_storage_filtered.csv -NoTypeInformation

    $ExcelParamsPhysical = @{
        WorksheetName = 'volume_detail'
        TableName = 'volume_detail'
        TableStyle = 'Medium2'
        AutoSize = $true
    }

    $ExportVolumeDetail | Export-Excel $Report @ExcelParamsPhysical
}
Get-VolumeDetail

Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $From -To $To -Subject $Subject -Body $Body -Attachments $Report –DeliveryNotificationOption OnSuccess

Remove-Item $Report 
