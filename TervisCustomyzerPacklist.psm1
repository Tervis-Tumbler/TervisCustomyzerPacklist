$ModulePath = if ($PSScriptRoot) {
    $PSScriptRoot
} else {
    (Get-Module -ListAvailable TervisApplication).ModuleBase
}

function Invoke-CutomyzerPackListProcess {
	param (
		$EnvironmentName,
		[Parameter(Mandatory)]$SiteCodeID
	)
	$BatchNumber = New-CustomyzerPacklistBatch -SiteCodeID $SiteCodeID
	if ($BatchNumber) {
		$DateTime = Get-Date
		$DocumentFilePaths = Invoke-CustomyzerPackListDocumentsGenerate -BatchNumber $BatchNumber -DateTime $DateTime -EnvironmentName $EnvironmentName
		
		$DocumentFilePaths |
		Send-CustomyzerPackListDocument -EnvironmentName $EnvironmentName -DateTime $DateTime -BatchNumber $BatchNumber
		
		$DocumentFilePaths |
		Send-CustomyzerPackListDocumentToArchive -EnvironmentName $EnvironmentName
	}
}

function New-CustomyzerPacklistBatch {
	param (
		[Parameter(Mandatory)]$SiteCodeID
	)
	$PackListLinesNotInBatch = Get-CustomyzerApprovalPackList -NotInBatch -SiteCodeID $SiteCodeID

	if ($PackListLinesNotInBatch) {
		$BatchNumber = New-CustomyzerBatchNumber
		$PackListLinesNotInBatch | Set-CustomyzerApprovalPackList -BatchNumber $BatchNumber -SentDateUTC (Get-Date).ToUniversalTime()
		$BatchNumber
	}
}

function Invoke-CustomyzerPackListDocumentsGenerate {
	param (
		[Parameter(Mandatory)]$BatchNumber,
		$DateTime = (Get-Date),
		[Parameter(Mandatory)]$EnvironmentName
	)
	$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber
	$PackListLinesSorted = Invoke-CustomyzerPackListLinesSort -PackListLines $PackListLines
	$CustomyzerPackListTemporaryFolder = New-CustomyzerPackListTemporaryFolder -BatchNumber $BatchNumber -EnvironmentName $EnvironmentName

	$Parameters = @{
		BatchNumber = $BatchNumber
		PackListLines = $PackListLinesSorted
		Path = $CustomyzerPackListTemporaryFolder.FullName
	}

	[PSCustomObject]@{
		XLSXFilePath = New-CustomyzerPacklistXlsx @Parameters -DateTime $DateTime
		CSVFilePath = New-CustomyzerPurchaseRequisitionCSV @Parameters
		XMLFilePath = New-CustomyzerPackListXML @Parameters -DateTime $DateTime
		XMLRewriteFinalArchedImageLocationForNewWebToPrintFilePath = New-CustomyzerPackListXML @Parameters -DateTime $DateTime -RewriteFinalArchedImageLocationForNewWebToPrint
	}
}

function New-CustomyzerPackListTemporaryFolder {
	param (
		[Parameter(Mandatory)]$BatchNumber,
		[Parameter(Mandatory)]$EnvironmentName
	)
	$TemporaryFolderPath = "$([System.IO.Path]::GetTempPath())\$EnvironmentName-$BatchNumber"
	Remove-Item -LiteralPath $TemporaryFolderPath -Force -Recurse -ErrorAction SilentlyContinue
	New-Item -ItemType Directory -Path $TemporaryFolderPath -Force
}

function Invoke-CustomyzerPackListLinesSort {
	param (
		[Parameter(Mandatory)]$PackListLines
	)
	$PackListLines |
	Sort-Object -Property {$_.OrderDetail.Project.Product.Form.Size},
		{$_.SizeAndFormType},
		{$_.OrderDetail.Order.ERPOrderNumber},
		{$_.OrderDetail.ERPOrderLineNumber}
}

function New-CustomyzerPacklistXlsx {
	param (
		[Parameter(Mandatory)]$BatchNumber,
		[Parameter(Mandatory)]$PackListLines,
		[Parameter(Mandatory)]$Path,
		$DateTime = (Get-Date)
	)

	$RecordToWriteToExcel = foreach ($PackListLine in $PackListLines) {
		[PSCustomObject]@{
			FormSize = $PackListLine.SizeAndFormType
			SalesOrderNumber = $PackListLine.OrderDetail.Order.ERPOrderNumber
			DesignNumber = $PackListLine.OrderDetail.ERPOrderLineNumber
			BatchNumber = $PackListLine.BatchNumber
			Quantity = $PackListLine.Quantity
			ScheduleNumber = $PackListLine.ScheduleNumber
		}
	}

	$Excel = Export-Excel -Path $Script:ModulePath\PackListTemplate.xlsx -PassThru
	$PackingListWorkSheet = $Excel.Workbook.Worksheets["PackingList"]

	Set-CustomyzerPackListXlsxHeaderValues -PackingListWorkSheet $PackingListWorkSheet -PackListXlsxLines $RecordToWriteToExcel -DateTime $DateTime
	Set-CustomyzerPackListXlsxRowValues -PackingListWorkSheet $PackingListWorkSheet -PackListXlsxLines $RecordToWriteToExcel

	$XlsxFileName = "TervisPackList-$BatchNumber.xlsx"
	$XlsxFilePath = "$Path\$XlsxFileName"
	$Excel.SaveAs($XlsxFilePath)
	$XlsxFilePath 
}

function Set-CustomyzerPackListXlsxHeaderValues {
	param (
		[Parameter(Mandatory)]$PackingListWorkSheet,
		[Parameter(Mandatory)]$PackListXlsxLines,
		$DateTime = (Get-Date)
	)
	$PackingListWorkSheet.Names["DownloadDate"].value = $DateTime.ToString("MM/dd/yyyy")
	$PackingListWorkSheet.Names["DownloadTime"].value = $DateTime.ToString("hh:mm tt")

	$FormSizeQuantitySums = $PackListXlsxLines |
	Group-Object FormSize |
	Add-Member -MemberType ScriptProperty -Name QuantitySum -PassThru -Force -Value {
		$This.Group |
		Measure-Object -Property Quantity -Sum |
		Select-Object -ExpandProperty Sum
	} |
	Add-Member -MemberType ScriptProperty -Name TotalName -PassThru -Force -Value {
		"Total$($This.Name)"
	} |
	Select-Object -Property QuantitySum, TotalName

	foreach ($FormSizeQuantitySum in $FormSizeQuantitySums) {
		$PackingListWorkSheet.Names[$FormSizeQuantitySum.TotalName].value = $FormSizeQuantitySum.QuantitySum
	}

	$GrandTotal = $FormSizeQuantitySums | Measure-Object -Property QuantitySum -Sum | Select-Object -ExpandProperty Sum
	$PackingListWorkSheet.Names["GrandTotal"].Value = $GrandTotal
}

function Set-CustomyzerPackListXlsxRowValues {
	param (
		[Parameter(Mandatory)]$PackingListWorkSheet,
		[Parameter(Mandatory)]$PackListXlsxLines
	)
	[int]$StartOfPackListLineDataRowNumber = $PackingListWorkSheet.Names["FirstCellOfPacklistLineData"].FullAddressAbsolute -split "\$" |
	Select-Object -Last 1

	$PropertyNameToColumnLetterMap = @{
		FormSize = "A"
		SalesOrderNumber = "C"
		DesignNumber = "E"
		BatchNumber = "G"
		Quantity = "I"
		ScheduleNumber = "P"
	}

	$PackListXlsxLines |
	ForEach-Object -Begin {
		$PackListXlsxLinesIndexNumber = 0
	} -Process {
		$Line = $_
		$RowNumber = $PackListXlsxLinesIndexNumber + $StartOfPackListLineDataRowNumber
		$PropertyNames = $Line.psobject.Properties.name

		foreach ($PropertyName in $PropertyNames) {
			$ColumnLetter = $PropertyNameToColumnLetterMap.$PropertyName
			$CellAddress = "$ColumnLetter$RowNumber"
			$PackingListWorkSheet.Cells[$CellAddress].Value = $Line.$PropertyName
		}

		$PackListXlsxLinesIndexNumber += 1
	}
}
function Set-StringValueFirstOccurence {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$String,
        [Parameter(Mandatory)]$OldValue,
        [Parameter(Mandatory)]$NewValue
    )
    process {
        $PositionOfOldValueInString = $String.IndexOf($OldValue)
        if ($PositionOfOldValueInString -ne -1) {
            $String.Substring(0, $PositionOfOldValueInString) + $NewValue + $String.Substring($PositionOfOldValueInString + $OldValue.Length)
        } else {
            $String
        }    
    }
}

function New-CustomyzerPackListXML {
	param (
		[Parameter(Mandatory)]$BatchNumber,
		[Parameter(Mandatory)]$PackListLines,
		[Parameter(Mandatory)]$Path,
		$DateTime = (Get-Date),
		[Switch]$RewriteFinalArchedImageLocationForNewWebToPrint
	)
	$InsertAfterQuantity = 500

	$XMLContent = New-XMLDocument -AsString -InnerElements {
		New-XMLElement -Name packList -InnerElements {
			New-XMLElement -Name batchNumber -InnerText $BatchNumber
			New-XMLElement -Name batchDate -InnerText $DateTime.ToString("MM/dd/yyyy")
			New-XMLElement -Name batchTime -InnerText $DateTime.ToString("hh:mm tt")
			New-XMLElement -Name orders -InnerElements {
				foreach ($PackListLine in $PackListLines) {
					New-TervisCustomyzerPackListOrderXmlElement `
						-PackListLine $PackListLine `
						-ItemQuantity $NumberOfInsert * $InsertAfterQuantity `
						-FileNameCDATAValue $(
							New-CustomyzerPackListSeparatorWrapPrintableFileURL `
								-ProductFormType $PackListLine.OrderDetail.Project.Product.Form.FormType.ToUpper() `
								-ProductSize $PackListLine.OrderDetail.Project.Product.Form.Size `
								-SalesOrderNumber $PackListLine.OrderDetail.Order.ERPOrderNumber `
								-SalesLineNumber $PackListLine.OrderDetail.ERPOrderLineNumber `
								-ScheduleNumber $PackListLine.ScheduleNumber
						)
					New-TervisCustomyzerPackListOrderXmlElement `
						-PackListLine $PackListLine `
						-RewriteFinalArchedImageLocationForNewWebToPrint:$RewriteFinalArchedImageLocationForNewWebToPrint
				}
			}
		}
	}

	$XMLFileName = if (-not $RewriteFinalArchedImageLocationForNewWebToPrint) {
		"TervisPackList-$BatchNumber-OldWebToPrint.xml2"
	} elseif ($RewriteFinalArchedImageLocationForNewWebToPrint) {
		"TervisPackList-$BatchNumber-NewWebToPrint.xml"
	}

	$XMLFilePath = "$Path\$XMLFileName"
	$XMLContent | Out-File -FilePath $XMLFilePath -Encoding ascii
	$XMLFilePath
}

function New-TervisCustomyzerPackListOrderXmlElement {
	param (
		[Parameter(Mandatory)]$PackListLine,
		$ItemQuantity,
		[Parameter(Mandatory)]$FileNameCDATAValue,
		[Switch]$RewriteFinalArchedImageLocationForNewWebToPrint
	)

	if (-not $ItemQuantity) {
		$ItemQuantity = $PackListLine.Quantity
	}

	if (-not $FileNameCDATAValue) {
		$FileNameCDATAValue = if (-not $RewriteFinalArchedImageLocationForNewWebToPrint) {
			$PackListLine.OrderDetail.Project.FinalArchedImageLocation	
		} elseif ($RewriteFinalArchedImageLocationForNewWebToPrint) {
			$PackListLine.OrderDetail.Project.FinalArchedImageLocation |
			Set-StringValueFirstOccurence -OldValue "http://images.tervis.com" -NewValue "http://images2.tervis.com"	
		}
	}

	New-XMLElement -Name order -InnerElements {
		New-XMLElement -Name salesOrderNumber -InnerText $PackListLine.OrderDetail.Order.ERPOrderNumber
		New-XMLElement -Name salesLineNumber -InnerText $PackListLine.OrderDetail.ERPOrderLineNumber
		New-XMLElement -Name itemQuantity -InnerText $ItemQuantity
		New-XMLElement -Name size -InnerText $PackListLine.SizeAndFormType
		New-XMLElement -Name itemNumber -InnerText $PackListLine.OrderDetail.Project.FinalFGCode
		New-XMLElement -Name scheduleNumber -InnerText $PackListLine.ScheduleNumber
		New-XMLElement -Name fileName -InnerElements {
			New-XMLCDATA -Value $(
				$FileNameCDATAValue
			)
		}
	}
}

function New-CustomyzerPackListSeparatorWrapPrintableFileURL {
	param (
		$SalesOrderNumber,
		$SalesLineNumber,
		$ScheduleNumber,
		$ProductSize,
		$ProductFormType
	)
	$ProductMetaData = Get-CustomyzerSizeAndFormTypeMetaData -Size $ProductSize -FormType $ProductFormType

	$TextValue = (@(
		$SalesOrderNumber,
		$SalesLineNumber,
		$ScheduleNumber,
		$ProductSize,
		$ProductFormType
	) | ForEach-Object { "\fs200%20$_\line"}) -join ""
	$ColorInkImageURL = [System.Web.HttpUtility]::UrlEncode("https://images.tervis.com/is/image/tervis?wid=$($ProductMetaData.PrintImageDimensions.Width)&hei=$($ProductMetaData.PrintImageDimensions.Height)&text=$TextValue&fmt=png-alpha&scl=1")
	$WhiteInkImageURL = [System.Web.HttpUtility]::UrlEncode("https://images.tervis.com/is/image/tervis?wid=$($ProductMetaData.PrintImageDimensions.Width)&hei=$($ProductMetaData.PrintImageDimensions.Height)&text=$TextValue&op_invert=1&fmt=png,gray&scl=1")
	@"
		https://images2.tervis.com/PrintableFile?
		ColorInkImageURL=$ColorInkImageURL
		&WhiteInkImageURL=$WhiteInkImageURL
		&OrderNumber=$SalesOrderNumber/$SalesLineNumber
		&ProductSize=$ProductSize
		&ProductFormType=$ProductFormType
"@ | Remove-WhiteSpace
}

function New-CustomyzerPurchaseRequisitionCSV {
	param (
		[Parameter(Mandatory)]$BatchNumber,
		[Parameter(Mandatory)]$PackListLines,
		[Parameter(Mandatory)]$Path
	)

	$RecordToWriteToCSV = foreach ($PackListLine in $PackListLines) {
		[PSCustomObject]@{
			ITEM_NUMBER = $PackListLine.OrderDetail.Project.FinalFGCode
			INTERFACE_SOURCECODE = "MIZER_REQ_IMPORT"
			SALES_ORDER_NO = $PackListLine.OrderDetail.Order.ERPOrderNumber
			SO_LINE_NO =$PackListLine.OrderDetail.ERPOrderLineNumber
			QUANTITY = $PackListLine.Quantity
			VENDOR_BATCH_NAME = $PackListLine.BatchNumber
			SCHEDULE_NUMBER = $PackListLine.ScheduleNumber
		}
	}
	#This should be the simple way to accomplish what is needed but we need to confirm that whatever
	#consumes this CSV can handle the values between the delimiter being in "" as currently they are not
	#$RecordToWriteToCSV |
	#ConvertTo-Csv -Delimiter "|" -NoTypeInformation |
	#Out-File -Encoding ascii -FilePath .\$CSVFileName -Force
	
	$CSVHeader = "ITEM_NUMBER",
	"INTERFACE_SOURCECODE",
	"SALES_ORDER_NO",
	"SO_LINE_NO",
	"QUANTITY",
	"VENDOR_BATCH_NAME",
	"SCHEDULE_NUMBER" -join "|"
	
	$CSVRows = $RecordToWriteToCSV |
	ForEach-Object { $_.psobject.Properties.value -join "|" }
	
	$CSVFileName = "xxmizer_reqimport_$BatchNumber.csv"
	$CSVFilePath = "$Path\$CSVFileName"
	$CSVHeader,$CSVRows | Out-File -Encoding ascii -FilePath $CSVFilePath
	$CSVFilePath
}

function Send-CustomyzerPackListDocument {
	param (
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XLSXFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XMLFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XMLRewriteFinalArchedImageLocationForNewWebToPrintFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$CSVFilePath,
		[Parameter(Mandatory)]$EnvironmentName,
		[Parameter(Mandatory)]$BatchNumber,
		$DateTime = (Get-Date)
	)
	process {
		$CustomyzerEnvironment = Get-CustomyzerEnvironment -EnvironmentName $EnvironmentName

		$MailMessageParameters = @{
			From = "customercare@tervis.com"
			To = $CustomyzerEnvironment.EmailAddressToRecieveXLSX
			Subject = "$($CustomyzerEnvironment.Name) Packlist"
			Attachments = $XLSXFilePath
			Body =  @"
<p>Packlist generated for batch - $BatchNumber</p>
<p><b>Created Date: </b>$($DateTime.ToString("MM/dd/yyyy"))</p>
<p><b>Created Time: </b>$($DateTime.ToString("hh:mm tt"))</p>
"@
		}
		Send-TervisMailMessage @MailMessageParameters -BodyAsHTML

		New-PSDrive -Name PackListXMLDestination -PSProvider FileSystem -Root $CustomyzerEnvironment.PackListXMLDestinationPath -Credential $CustomyzerEnvironment.FileShareAccount | Out-Null
		$XMLFilePath, $XMLRewriteFinalArchedImageLocationForNewWebToPrintFilePath, $XLSXFilePath | 
		Copy-Item -Destination PackListXMLDestination:\ -Force
		Remove-PSDrive -Name PackListXMLDestination

		Set-TervisEBSEnvironment -Name $EnvironmentName
		$EBSIASNode = Get-EBSIASNode
		Set-SFTPFile -RemotePath $CustomyzerEnvironment.RequisitionDestinationPath -LocalFile $CSVFilePath -SFTPSession $EBSIASNode.SFTPSession -Overwrite:$Overwrite
	}
}

function Send-CustomyzerPackListDocumentToArchive {
	param (
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XLSXFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XMLFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$XMLRewriteFinalArchedImageLocationForNewWebToPrintFilePath,
		[Parameter(Mandatory,ValueFromPipelineByPropertyName)]$CSVFilePath,
		[Parameter(Mandatory)]$EnvironmentName
	)
	process {
		$CustomyzerEnvironment = Get-CustomyzerEnvironment -EnvironmentName $EnvironmentName
		$ArchivePath = "$($CustomyzerEnvironment.PackListFilesPathRoot)\Inbound\PackLists\Archive"

		New-PSDrive -Name Archive -PSProvider FileSystem -Root $ArchivePath -Credential $CustomyzerEnvironment.FileShareAccount | Out-Null
		$XLSXFilePath, $XMLFilePath, $XMLRewriteFinalArchedImageLocationForNewWebToPrintFilePath, $CSVFilePath |
		Copy-Item -Destination Archive:\ -Force
		Remove-PSDrive -Name Archive
	}
}

function Install-CustomyzerPackListGenerationApplication {
	param (
		[Parameter(Mandatory)]$EnvironmentName,
		[Parameter(Mandatory)]$ComputerName
	)
	$Environment = Get-CustomyzerEnvironment -EnvironmentName $EnvironmentName
	$PasswordstateAPIKey = Get-TervisPasswordstatePassword -Guid $Environment.PasswordStateAPIKeyPasswordGUID |
	Select-Object -ExpandProperty Password
	
	$PowerShellApplicationParameters = @{
		ComputerName = $ComputerName
		EnvironmentName = $EnvironmentName
		ModuleName = "TervisCustomyzer"
		RepetitionIntervalName = $Environment.ScheduledTaskRepetitionIntervalName
		ScheduledTasksCredential = New-Crednetial -Username system
		ScheduledTaskName = "CustomyzerPackListGeneration"
		TervisModuleDependencies = @"
TervisMicrosoft.PowerShell.Security
TervisMicrosoft.PowerShell.Utility
PasswordstatePowerShell
TervisPasswordstatePowerShell
TervisApplication
WebServicesPowerShellProxyBuilder
InvokeSQL
PowerShellORM
OracleE-BusinessSuitePowerShell
TervisMailMessage
TervisOracleE-BusinessSuitePowerShell
TervisCustomyzer
TervisCustomyzerPacklist
"@ -split "`r`n"
	PowerShellGalleryDependencies = @"
ImportExcel
posh-ssh
"@ -split "`r`n"
	CommandString = @"
Set-PasswordstateAPIKey -APIKey $PasswordstateAPIKey
Set-PasswordstateAPIType -APIType Standard
Set-CustomyzerModuleEnvironment -Name $EnvironmentName
Invoke-CutomyzerPackListProcess -EnvironmentName $EnvironmentName -SiteCodeID 3
"@
	}

	Install-PowerShellApplication @PowerShellApplicationParameters
}

function New-CustomizerPackListXMLWithJust16OzOneOff {
    $EnvironmentName = "Production"
	$DateTime = (Get-Date)

	Set-CustomyzerModuleEnvironment -Name $EnvironmentName
	# Get-CustomyzerApprovalPacklistRecentBatch

	$BatchNumber = "20190109-1300"
	foreach ($BatchNumber in $BatchNumbers) {
		$PackListLines = Get-CustomyzerApprovalPackList -BatchNumber $BatchNumber
		$PackListLinesSorted = Invoke-CustomyzerPackListLinesSort -PackListLines $PackListLines
		$Size16PackListLinesSorted = $PackListLinesSorted |
		Where-Object {
			$_.Orderdetail.Project.Product.Form.Size -eq 16
		}

		$CustomyzerPackListTemporaryFolder = New-CustomyzerPackListTemporaryFolder -BatchNumber $BatchNumber -EnvironmentName $EnvironmentName

		$Parameters = @{
			BatchNumber = $BatchNumber
			PackListLines = $PackListLinesSorted
			Path = $CustomyzerPackListTemporaryFolder.FullName
		}
		$XMLFilePath = New-CustomyzerPackListXML @Parameters -DateTime $DateTime
	}

	# $MissingOne = $Size16PackListLinesSorted | where {
	#     -not $_.OrderDetail.Project.FinalArchedImageLocation
	# }
}

function Test-PDFFilesForCorruption {
	param (
		$Path
	)
	$PDFFiles = Get-ChildItem -Path $Path -Recurse -Filter *.pdf
	$Jobs = foreach ($PDFFile in $PDFFiles) {
		Start-RSJob -ScriptBlock {
			Set-Location -Path $($Using:PDFFile).Directory
			wsl pdfinfo $($Using:PDFFile).Name | Out-Null
			if ($LastExitCode) {
				$($Using:PDFFile).Name
			}
		}	
	}

	$Jobs |
	Wait-RSJob |
	Receive-RSJob

	$Jobs |
	Remove-RSJob
}

function Invoke-CustomyzerPacklistDownload {
	param (
		$XmlFilePath,
		$PDFOutputFolderPath
	)
	$Content = Get-Content -Path $XmlFilePath
	[XML]$PackListXMLLines = $Content
	$FolderName = $PackListXMLLines.packList.batchNumber #Ex: 20190424-1300
	
	$PackListXMLLines.packList.orders.order |
	Add-Member -MemberType ScriptProperty -Name PDFFileName -Value {
		"$($This.Size)-$($This.salesOrderNumber)-$($This.salesLineNumber)-$($This.itemQuantity).pdf" #Ex: 16DWT-11488993-6-1.pdf
	} -Force -PassThru |
	Add-Member -MemberType ScriptProperty -Name OutFilePath -Value {
		"$PDFOutputFolder\$FolderName\$($This.PDFFileName)"
	}

	$PackListXMLLines.packList.orders.order |
	ForEach-Object {
		$PackListLine = $_
		Start-ThreadJob -ThrottleLimit 10 -ScriptBlock {
			$PackListLine = $Using:PackListLine
			$OutFilePath = "$Using:PDFOutputFolder\$Using:FolderName\$($PackListLine.PDFFileName)"
			if (-not (Test-Path -Path $OutFilePath)) {
				$ProgressPreference = "SilentlyContinue"
		
				$URL = "$($PackListLine.Filename.'#cdata-section')&`$orderNum=$($PackListLine.salesOrderNumber)/$($PackListLine.salesLineNumber)"
				Invoke-WebRequest -OutFile $OutFilePath -Uri $URL -UseBasicParsing
			}
		}
	} |
	Receive-Job -AutoRemoveJob -Wait
}