$Content = Get-Content -Path "C:\Users\c.magnuson\OneDrive - tervis\Downloads\Packlist download testing\TervisPackList-20190425-1301-NewWebToPrint.xml"
[XML]$PackListXMLLines = $Content
$PDFOutputFolder = "C:\Users\c.magnuson\OneDrive - tervis\Downloads\Packlist download testing\PDFs"
#20190424-1300
$FolderName = $PackListXMLLines.packList.batchNumber
$PackListXMLLines.packList.orders.order |
Add-Member -MemberType ScriptProperty -Name PDFFileName -Value {
    #16DWT-11488993-6-1.pdf
    "$($This.Size)-$($This.salesOrderNumber)-$($This.salesLineNumber)-$($This.itemQuantity).pdf"
} -Force -PassThru |
Add-Member -MemberType ScriptProperty -Name OutFilePath -Value {
    "$PDFOutputFolder\$FolderName\$($This.PDFFileName)"
}

#$PackListXMLLines.packList.orders.order | Format-Table
#$PackListLine = ($PackListXMLLines.packList.orders.order| Select-Object -First 1)

$PackListXMLLines.packList.orders.order |
ForEach-Object {
    $PackListLine = $_
    if ($false) {
        $OutFilePath = 
        if (-not (Test-Path -Path $OutFilePath)) {
            Write-Verbose -Message "Downloading $OutFilePath Start"
            $ProgressPreference = "SilentlyContinue"
    
            $URL = "$($PackListLine.Filename.'#cdata-section')&`$orderNum=$($PackListLine.salesOrderNumber)/$($PackListLine.salesLineNumber)"
            Invoke-WebRequest -OutFile $OutFilePath -Uri $URL -UseBasicParsing
            Write-Verbose -Message "Downloading $OutFilePath Finish"
        } else {
            Write-Verbose "$OutFilePath already exists, skipping"
        }  
    } else {
        Start-ThreadJob -ThrottleLimit 10 -ScriptBlock {
            $PackListLine = $Using:PackListLine
            $OutFilePath = "$Using:PDFOutputFolder\$Using:FolderName\$($PackListLine.PDFFileName)"
            if (-not (Test-Path -Path $OutFilePath)) {
                Write-Output -Message "Downloading $OutFilePath Start"
                $ProgressPreference = "SilentlyContinue"
        
                $URL = "$($PackListLine.Filename.'#cdata-section')&`$orderNum=$($PackListLine.salesOrderNumber)/$($PackListLine.salesLineNumber)"
                Write-Output $URL
                Invoke-WebRequest -OutFile $OutFilePath -Uri $URL -UseBasicParsing
                Write-Output -Message "Downloading $OutFilePath Finish"
            } else {
                Write-Output "$OutFilePath already exists, skipping"
            }  
        }
    }
} |
Receive-Job -AutoRemoveJob -Wait


$GraphicsBatchRenderLogLinesToFixQuery = @"
SELECT PartitionedTable.*
FROM (  
	SELECT *,
		ROW_NUMBER() OVER(
			PARTITION BY [GraphicsBatchHeaderID],
				[BatchRenderType],
				[SalesOrderNumber],
				[SalesLineNumber]
			ORDER BY [ProcessingFinish] desc
		) as RowNumber
    FROM [MES].[Graphics].[GraphicsBatchRenderLog] with (nolock)
	WHERE GraphicsBatchHeaderID > 19996
) as PartitionedTable
WHERE PartitionedTable.RowNumber = 1
AND ProcessingStatus = 'Failed'
order by id desc
"@

$GraphicsBatchRenderLogLinesToFix = Invoke-MSSQL -Server SQL -Database MES -SQLCommand $GraphicsBatchRenderLogLinesToFixQuery -ConvertFromDataRow

# $GraphicsBatchRenderLogLinesToFix = Get-MESGraphicsGraphicsBatchRenderLog -GraphicsBatchHeaderID 20018

$GraphicsBatchRenderLogLinesToFix |
Add-Member -MemberType ScriptProperty -Name SourceURLNewWebToPrint -Value {
    $This.SourceURL |
    Set-StringValueFirstOccurence -OldValue "http://images.tervis.com" -NewValue "http://images2.tervis.com"
} -Force  -PassThru |
Add-Member -MemberType ScriptProperty -Name DestinationURLEscaped -Value {
    [WildcardPattern]::Escape($This.DestinationUrl)
} -Force -PassThru |
Add-Member -MemberType ScriptProperty -Name TempFileName -Value {
    "$($This.DestinationUrl).tmp"
} -Force

#Where-Object GraphicsBatchHeaderID -eq 20080 |

$GraphicsBatchRenderLogLinesToFix |
ForEach-Object {
    $GraphicsBatchRenderLogLine = $_
    if (-not (Test-Path -LiteralPath $GraphicsBatchRenderLogLine.DestinationURL)) {
    Start-ThreadJob -ThrottleLimit 10 -ScriptBlock {
        $ErrorActionPreference = "Inquire"
            $ProgressPreference = "SilentlyContinue"
            $GraphicsBatchRenderLogLine = $Using:GraphicsBatchRenderLogLine
            Write-Verbose $GraphicsBatchRenderLogLine.DestinationURL
            $GraphicsBatchRenderLogLine | Set-MESGraphicsGraphicsBatchRenderLog -ProcessingStatus Processing
            Invoke-WebRequest -Uri $GraphicsBatchRenderLogLine.SourceURLNewWebToPrint -OutFile $($GraphicsBatchRenderLogLine.DestinationURLEscaped) -UseBasicParsing
            Rename-Item -LiteralPath $GraphicsBatchRenderLogLine.DestinationURLEscaped -NewName $GraphicsBatchRenderLogLine.DestinationURL
            Remove-Item -LiteralPath $GraphicsBatchRenderLogLine.TempFileName -ErrorAction SilentlyContinue
            $GraphicsBatchRenderLogLine | Set-MESGraphicsGraphicsBatchRenderLog -ProcessingStatus Success
        }
    }
} |
Receive-Job -AutoRemoveJob -Wait

  
        

# public enum GraphicsProcessingStatus
# {
#     Staged,
#     Processing,
#     //Cancelled,
#     Failed,
#     Success,
# }

# update [Graphics].[GraphicsBatchRenderLog]
# set
# [ProcessingStatus] = 'Success'

# where 1 = 1
# AND [Graphics].[GraphicsBatchRenderLog].[ID] in (
# '246030',
# '245941'
# )

#UpdateBatchRenderLogStatus http://tfs2012:8080/tfs/DefaultCollection/MES/_search?type=Code&lp=search-project&text=def%3AUpdateBatchRenderLogStatus&result=DefaultCollection%2FMES%2F%24%2FMES%2F7968%2F%24%2FMES%2FSource%2FDev%2FTervis.MES.Graphics%2FTervis.MES.Graphics.Shared.Data%2FRepository%2FBatchRepository.cs&filters=ProjectFilters%7BMES%7D&preview=1&_a=contents
#WebClient_DownloadProgressChanged http://tfs2012:8080/tfs/DefaultCollection/MES/_search?type=Code&lp=search-project&text=def%3ADownloadBatchFile&result=DefaultCollection%2FMES%2F%24%2FMES%2F8021%2F%24%2FMES%2FSource%2FMain%2FTervis.MES.Graphics%2FTervis.MES.Graphics.Services%2FCustomizerService.cs&filters=ProjectFilters%7BMES%7D&preview=1&_a=contents
#DownloadHelixPdfFilesAsync