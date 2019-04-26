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
    FROM [MES].[Graphics].[GraphicsBatchRenderLog]
	WHERE GraphicsBatchHeaderID > 19996
	AND ProcessingStatus = 'Failed'
) as PartitionedTable
WHERE PartitionedTable.RowNumber = 1
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
} -Force -PassThru |
Where-Object GraphicsBatchHeaderID -eq 20080 |
#Select-Object -First 30 |
ForEach-Object {
    $GraphicsBatchRenderLogLine = $_
    if (-not (Test-Path -LiteralPath $GraphicsBatchRenderLogLine.DestinationURL)) {
        try {
            $GraphicsBatchRenderLogLine | Set-MESGraphicsGraphicsBatchRenderLog -ProcessingStatus Processing
            Invoke-WebRequest -Uri $GraphicsBatchRenderLogLine.SourceURLNewWebToPrint -OutFile $($GraphicsBatchRenderLogLine.DestinationURLEscaped) -UseBasicParsing
            Rename-Item -LiteralPath $GraphicsBatchRenderLogLine.DestinationURLEscaped -NewName $GraphicsBatchRenderLogLine.DestinationURL
            Remove-Item -LiteralPath $GraphicsBatchRenderLogLine.TempFileName
            $GraphicsBatchRenderLogLine | Set-MESGraphicsGraphicsBatchRenderLog -ProcessingStatus Success    
        } catch {

        }
    }
}

# public enum GraphicsProcessingStatus
# {
#     Staged,
#     Processing,
#     //Cancelled,
#     Failed,
#     Success,
# }