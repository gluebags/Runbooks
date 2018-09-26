#ParsePDF-DEV RunBook
#github@cityrisk.net 
#Updated 27/09/18

#Get variables stored in Runbook
$storageAccountName = Get-AutomationVariable -Name "StorageAccount-Name"
$storageAccessKey = Get-AutomationVariable -Name "StorageAccount-Key"			
$ContainerName = Get-AutomationVariable -Name "StorageContainer-Name"
$localFileDirectory = Get-AutomationVariable -Name "LocalDirectory-Path"
$BlobName = Get-AutomationVariable -Name "BlobName"
$localFile = Get-AutomationVariable -Name "LocalFile-Path"
$amountmationAccount = Get-AutomationVariable -Name "AutomationAccount-Name"
$resourceGroup = Get-AutomationVariable -Name "ResourceGroup-Name"
$outBoundWebhook = Get-AutomationVariable -Name "Outbound-JSON"

#Create Storage Context to access blob
$ctx = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccessKey

# Load itextsharp Module
Add-Type -Path "C:\modules\user\itextsharp\itextsharp.dll"

#Get File from Blob
$null = Get-AzureStorageBlobContent -Destination $localFileDirectory -Container $ContainerName -Blob $BlobName -Context $ctx

#Call our functions
Write-Output "Executing: Get-Calculations -Path $localfile"
$billDetails = Get-Calculations -Path $localFile

#convert output to json
$json = $billDetails | ConvertTo-json

#show me it worked
Write-Output $json

#send it
$uri = $outBoundWebhook
Invoke-RestMethod -Uri $uri `
                    -Method Post `
                    -Body $json `
                    -ContentType 'application/json'

#clean the blob
$null = Remove-AzureStorageBlob -Container $ContainerName -Blob "Bill.pdf" -Context $ctx
#clean local system
Remove-Item $localFile

#Converts PDF using itextsharp
function Get-Conversion
	{
	[CmdletBinding()]
	[OutputType([string])]
	param (
	  [Parameter(Mandatory = $true)]
	  [string]
	  $Path
	)

	$Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

	try
	{
	  $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $Path
	}
	catch
	{
	  throw
	}

	$stringBuilder = New-Object System.Text.StringBuilder

	for ($page = 1; $page -le $reader.NumberOfPages; $page++)
	{
	  $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page)
	  $null = $stringBuilder.AppendLine($text) 
	}

	$reader.Close()

	return $stringBuilder.ToString()
	}

function Get-Calculations {
	param (
		[string]$Path
	)

	# Convert Bill to text
	$tmp = Get-Conversion -Path $localFile
	# Get the dollar amount due for this bill
	$amount = [Regex]::Match($tmp, "(Due \$) (\d+\.\d+)").Captures.groups[2].Value

	# Work out individual costs
	$individualCost = ($amount/5)

	# Pass today's and future date to our json object
	$date = Get-Date -date $(Get-Date) -Format (Get-Culture).DateTimeFormat.ShortDatePattern
	$dueDate = Get-Date -date $(Get-Date).AddDays(21) -Format (Get-Culture).DateTimeFormat.ShortDatePattern

	#Make these varibles into strings to please Excel
	[string]$friendlyCost = $individualCost
	[string]$friendlyAmount = $amount

	# Make a PSObject home for the JSON
	$obj = New-Object -TypeName PSObject
	$obj | Add-Member -MemberType NoteProperty -Name amount -Value $friendlyAmount
	$obj | Add-Member -MemberType NoteProperty -Name Date_Today -Value $date
	$obj | Add-Member -MemberType NoteProperty -Name Date_Due -Value $dueDate
	$obj | Add-Member -MemberType NoteProperty -Name Individual_Cost -Value $friendlyCost

	  return $obj
	}
