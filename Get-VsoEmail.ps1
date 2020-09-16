<# 
.SYNOPSIS

Example/template to create an Outlook message from VSO (Azure DevOps/Team Foundation Services) queries using PowerShell.

.DESCRIPTION

Two functions are used (Get-TfsQueryItems and Get-TfsItemsAsTable) to process queries using the Azure DevOps REST API. Additonal Credential Manager and EnhancedHTML2 modules are used for retrieving Azure DevOps Personal Access Token (PAT) and email formatting respectively.

.PARAMETER Shift

Example use case is for sending Shift Hand-Off emails based on VSO queries.

.EXAMPLE

PS> .\Get-VsoEmail.ps1 -Shift 1 

Create handoff email as member of the first shift.

.PARAMETER Output

Determines how the script will output results. OPTIONAL parameter, default is to output as an Outlook message. If set to (p) or (preview) results will only be displayed to PowerShell console.

.EXAMPLE

PS> .\Get-VsoEmail.ps1 -Shift 1 -Output preview

Preview shift handoff results in console as member of the first shift.

.INPUTS

None. You cannot pipe objects to Get-VsoEmail.ps1.

.OUTPUTS

Outlook MailItem object or Tables to standard out with query data.

.NOTES
  
Prerequisites:
    
 * Add Azure DevOps PAT to Windows Credential Manager

 * Install-Module -Name CredentialManager
     
 * Install-Module -Name EnhancedHTML2

.LINK

github.com/davotronic5000/PowerShell_Credential_Manager

.LINK

github.com/PowerShellOrg/EnhancedHTML2
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('1', '2', '3')]
    [String]$Shift,
 
    [Parameter(Mandatory = $false)]
    [ValidateSet('p', 'preview', 'e', 'email')]
    [String]$Output = 'e' # Default to Outlook email output
)

# Embedded CSS Here-String for example email formatting
$style = @"

table {
    font-family: Calibri,Tahoma;
    border-collapse: collapse;
    font-size: 11pt;
    width: 100%;
}

h2 {
    font-family: Calibri,Tahoma;
    font-size: 16pt;
    font-weight: normal;
    background-color: #505050;
    margin-bottom: 0;
    color: white;
}

th { 
    padding-top: 4px;
    padding-bottom: 4px;
    text-align: left;
    background-color: #505050;
    color: #dcdcdc;
}

table.QueryTables td {
	border-bottom: 1px solid #dcdcdc;
    padding: 0 4px;
}

.odd  { background-color: #ffffff; }

.even { background-color: #e6f5ff; }

// Hover does not work with Outlook html parser, if output is piped to an html file it is functional

// .odd:hover { background-color: #b0b0b0; } 

// .even:hover { background-color: #b0b0b0; }

.yellow { background-color: yellow; }

.red { 
    background-color: red; 
    color: white;
}

"@

# Authentication settings and DevOps project paths

try { $PAT = (Get-StoredCredential -Target 'GenericCredentialName').GetNetworkCredential().Password } # Pull PAT from Windows Credential Manager
catch { Write-host $_ }
$encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))
$TFSServerUrl = 'https://ServerURL/DefaultCollection'
$TFSProject = 'ProjectName'

# Table names for respective queries
$TName = 'Table Name'

# Using static column names as this lends to more easily changed column headings
# Can use additional query to pull/push given column names automatically
$ColNames = 'ID', 'Pickup Date', 'Priority', 'Assigned To', 'State', 'Title'

# TFS Work Item names for respective Column names
$WINames = 'id', 'Custom.PickupDate', 'Custom.Priority', 'System.AssignedTo', 'System.State', 'System.Title'

function Get-TfsQueryItems
{
    param ([String]$Wiql)
    try {
        # Write-Verbose "Optional troubleshooting message specific to POST attempt"
        $QueryResult = (Invoke-RestMethod -Method Post -ContentType "application/json" -Body $Wiql -Headers @{Authorization = "Basic $encodedPat"} -uri "$($TFSServerUrl)/$($TFSProject)/_apis/wit/wiql?api-version=2.1").workitems

        if (!$QueryResult) {
            $QueryResult
        }
        else {
            # Write-Verbose "Optional troubleshooting message specific to GET attempt"
            $WorkItems = (Invoke-RestMethod -Method Get -Headers @{Authorization = "Basic $encodedPat"} -uri "$($TFSServerUrl)/_apis/wit/WorkItems?ids=$($QueryResult.id -join ",")&api-version=2.1").value
            $WorkItems
        }
    }
    catch [System.Net.WebException] {
        Write-host $_
    }
    catch {
        Write-host $_
    }
}

function Get-TfsItemsAsTable
{
    [CmdletBinding()]
    param (
        [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true)]
        [PSCustomObject]$WorkItemsReturned,
        [String]$TblName, 
        [String[]]$ColNames, 
        [String[]]$WINames,
        [bool]$TitleAsLink = $false
    )

    process 
    {    
        if ($WorkItemsReturned)
        {
            $tbl = New-Object System.Data.DataTable $TblName
            $cols = @()
    
            # Add array of column names to table
            for ($i = 0; $i -lt $ColNames.Length; $i++)
            {
                $cols += New-Object System.Data.DataColumn $ColNames[$i]
                $tbl.Columns.Add($cols[$i])
            }
    
            # Add workitems for each row to table
            foreach ($workitem in $WorkItemsReturned)
            {
                $row = $tbl.NewRow()
                for ($i = 0; $i -lt $ColNames.Length; $i++)
                {
                    # '0' index of my example VSO query uses the 'ID' column which is at the root of the Work Item (no 'fields' attribute)
                    if ($i -eq 0) {
                        $row.$($ColNames[$i]) = $workitem.$($WINames[$i])
                    }

                    # All additional data for my example use case are under the Work Item's 'fields' attribute

                    # Example to use more easily modified Date ToString modifier for Date columns that may be in UTC (string) format
                    elseif ($ColNames[$i] -like '*Date') {

                        # Convert pulled UTC string to DateTime, maintain UTC (needed due to auto conversion to Local time), and format/ToString as specified
                        $formattedDateTime = (([DateTime]$workitem.fields.$($WINames[$i])).ToUniversalTime()).ToString('MM/dd/yyyy HH:mm')
                        $row.$($ColNames[$i]) = $formattedDateTime
                        
                    }

                    # Example to use Title column as URL to Work Item
                    elseif (($ColNames[$i] -eq 'Title') -and ($TitleAsLink -eq $true)) { 
                        $hrefID = $workitem.$($WINames[0])
                        $hrefTitle = $workitem.fields.$($WINames[$i])
                        $row.$($ColNames[$i]) = "<a href='WorkItemUrlPath/$hrefID'>$hrefTitle</a>"
                    }

                    # Default direct transfer of pulled Work Item field data into table
                    else {
                        $row.$($ColNames[$i]) = $workitem.fields.$($WINames[$i])
                    }
                }
                $tbl.Rows.Add($row)
            }
        }
        else {
            $tbl = 'No results match the query' # String to match typical web interface response
        }
        $tbl
    }
}

$timeNowLocal = Get-Date
$timeNowUtc = $timeNowLocal.ToUniversalTime()

Switch ($Shift) {
        '1' { 
            # Example using a First Shift staring at 1600 UTC and ending at 2400 UTC

            $StartOfShift = [DateTime]::Today.AddHours(16)
            $EndOfShift = [DateTime]::Today.AddHours(24)

            if (($timeNowUtc -ge $StartOfShift) -and ($timeNowUtc -lt $EndOfShift)) {
                $dateValue = '@Today'
            }
            else {
                $dateValue = '@Today-1'
            }
        }
        '2' { 
            # Second Shift time frame 
        }
        '3' { 
            # Third Shift time frame
        }
}

# Example WIQL statement used for table output
# Uses conditional dateValue based on UTC time to provide appropriate output based on script execution time

$WiqlQuery = "{
            'query': 
            `"select [System.Id], [Custom.PickupDate], [Custom.Priority], [System.AssignedTo], [System.State], [System.Title]
            from WorkItems where [System.TeamProject] = @project and 
                [System.WorkItemType] = 'WorkItemType' and 
                [Custom.PickupDate] >= $dateValue and 
            order by [Custom.Priority] asc`"
}"

# Preview query output to console only

if (($Output -eq 'p') -or ($Output -eq 'preview')) {
    $Time = Get-Date -Format HH:mm:ss
    Write-Host $Time '| Beginning shift handoff queries...'

    $TblFromQuery = Get-TfsQueryItems -Wiql $WiqlQuery | Get-TfsItemsAsTable -TblName $TName -ColNames $ColNames -WINames $WINames

    $Time = Get-Date -Format HH:mm:ss
    Write-Host $Time '| ...Finished queries'

    # If using more than 10 columns, listing each name by -Property is an option to overcome Format-Tables default max columns displayed as well as 
    # avoiding unecessary columns e.g. 'RowError', 'RowState', etc.
    Write-Output $TblFromQuery | Format-Table
}
else {

    # Default to create Outlook message based on query(ies)

    $Time = Get-Date -Format HH:mm:ss
    Write-Host $Time '| Beginning shift handoff queries...'

    $TblFromQuery = Get-TfsQueryItems -Wiql $WiqlQuery | Get-TfsItemsAsTable -TblName $TName -ColNames $ColNames -WINames $WINames -TitleAsLink $true

    $Time = Get-Date -Format HH:mm:ss
    Write-Host $Time '| ...Finished queries'

    $params = @{'OddRowCssClass'='odd';
                'EvenRowCssClass'='even';
                'TableCssClass'='QueryTables';
                'Properties'='ID',
                    'Pickup Date',
                    @{n='Priority'; 
                      e={$_.'Priority'};
                      css={if ($_.'Priority' -eq 'P0 - Emergency') { 'red' }
                           elseif ($_.'Priority' -eq 'P1 - Warning') { 'yellow' }}},
                    'Assigned To',
                    'State',
                    'Title'}

    $LineCt = $TblFromQuery | Measure-Object -Line
    $TableTitle = 'Table Title:  ' + $LineCt.Lines.ToString()
    $FragmentQuery = $TblFromQuery | ConvertTo-EnhancedHTMLFragment @params -PreContent "<h2>$TableTitle</h2>"
    
    $Greeting = "Greeting (can be html formatted)"
    
    $Closing = "Closing (can be html formatted)"
    
    # Example is using single query however, typically multiple query fragments would be joined into an array

    $fragments = @($FragmentQuery)
    $body = ConvertTo-EnhancedHTML -HTMLFragments $fragments -CssStyleSheet $style -PreContent $Greeting -PostContent $Closing
    
    # Create Outlook object's MailItem to display (will open and display created email if Outlook is not already running)

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = "Destination@Example.com"
    $Mail.Subject = "Subject"
    $Mail.HTMLBody = "$($body)"
    $Mail.Display()
}