<#
    .SYNOPSIS
    Extracts tables from an HTML page and returns them as an array of objects.

    .PARAMETER URL
    The URL of the HTML page to extract tables from.

    .PARAMETER TableNumber
    The number of the table to extract (starting from 0). If not specified, all tables will be extracted.

    .PARAMETER LocalFile
    A switch to indicate that the URL parameter is a local file path rather than a web URL

#>

function Get-HTMLTables {
    param(
        [Parameter(Mandatory)]
        [String] $URL,

        [Parameter(Mandatory = $false)]
        [int] $TableNumber,

        [Parameter(Mandatory = $false)]
        [boolean] $LocalFile
    
    )

    [System.Collections.Generic.List[PSObject]]$tablesArray = @()

    if ($LocalFile) {
        $html = New-Object -ComObject 'HTMLFile'
        $source = Get-Content -Path $URL -Raw
        $html.IHTMLDocument2_write($source)

        # html does not have ParseHTML because it already an HTMLDocumentClass
        # Cast in array in case of only one element
        $tables = @($html.getElementsByTagName('TABLE'))
    }
    else {
        $WebRequest = Invoke-WebRequest $URL -UseBasicParsing
        
        # Parse HTML manually using regex since COM objects don't work reliably in PowerShell 7
        $htmlContent = $WebRequest.Content
        
        # Extract all table elements using regex
        $tableMatches = [regex]::Matches($htmlContent, '<table[^>]*>.*?</table>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        # Create mock table objects for compatibility
        $tables = @()
        foreach ($tableMatch in $tableMatches) {
            $tableHtml = $tableMatch.Value
            # Create a simple object with Rows property containing the raw HTML
            $mockTable = [PSCustomObject]@{
                InnerHtml = $tableHtml
                Rows      = @()
            }
            
            # Extract rows using regex
            $rowMatches = [regex]::Matches($tableHtml, '<tr[^>]*>.*?</tr>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            foreach ($rowMatch in $rowMatches) {
                $rowHtml = $rowMatch.Value
                $mockRow = [PSCustomObject]@{
                    InnerHtml = $rowHtml
                    Cells     = @()
                }
                
                # Extract cells (th or td)
                $cellMatches = [regex]::Matches($rowHtml, '<(th|td)[^>]*>.*?</(th|td)>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline)
                
                foreach ($cellMatch in $cellMatches) {
                    $cellHtml = $cellMatch.Value
                    # Check if it's a header cell
                    $isHeader = $cellHtml -match '^<th'
                    
                    # Extract text content
                    $innerText = [regex]::Replace($cellHtml, '<[^>]+>', '') -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'
                    
                    $mockCell = [PSCustomObject]@{
                        tagName   = if ($isHeader) { 'TH' } else { 'TD' }
                        InnerText = $innerText.Trim()
                    }
                    
                    $mockRow.Cells += $mockCell
                }
                
                $mockTable.Rows += $mockRow
            }
            
            $tables += $mockTable
        }
    }

    ## Extract the tables out of the web request
    if ($TableNumber) {
        #$table = $tables[$TableNumber]
        # Cast in array because only one element
        $tables = @($tables[$TableNumber])
    }

    ## Go through all of the rows in the table
    $tableNumber = 0
    foreach ($table in $tables) {
        $titles = @()
        $rows = @($table.Rows)

        $tableNumber++

        foreach ($row in $rows) {
            $cells = @($row.Cells)

            ## If we've found a table header, remember its titles
            if ($cells[0].tagName -eq 'TH') {
                $titles = @($cells | ForEach-Object { ('' + $_.InnerText).Trim() })
                continue
            }

            ## If we haven't found any table headers, make up names "P1", "P2", etc.
            if (-not $titles) {
                $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
            }

            ## Now go through the cells in the the row. For each, try to find the
            ## title that represents that column and create a hashtable mapping those
            ## titles to content
            $resultObject = [Ordered] @{
                'TableNumber' = $tableNumber
            }

            for ($counter = 0; $counter -lt $cells.Count; $counter++) {
                $title = $titles[$counter]
                if (-not $title) { continue }  
                $resultObject[$title] = ('' + $cells[$counter].InnerText).Trim()
            }

            ## And finally cast that hashtable to a PSCustomObject and add to $array
            $tablesArray.Add([PSCustomObject] $resultObject)
        }
    }

    return $tablesArray
}