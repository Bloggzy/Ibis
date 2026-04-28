function Add-IbisLogLine {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogTextBox,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$Level = 'INFO',

        [string]$LogFilePath = $script:IbisCurrentLogFilePath
    )

    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $displayMessage = ConvertTo-IbisGuiDisplayText -Text $Message -StripAnsi
    $LogTextBox.AppendText("[$timestamp] $displayMessage`r`n")
    Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message $Message -Level $Level
}

function ConvertTo-IbisGuiDisplayText {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Text,

        [switch]$StripAnsi
    )

    if ($null -eq $Text) {
        return ''
    }

    $displayText = [string]$Text
    if ($StripAnsi) {
        $escape = [regex]::Escape([string][char]27)
        $displayText = [regex]::Replace($displayText, "$escape\[[0-9;?]*[ -/]*[@-~]", '')
    }

    $displayText -replace "`r`n|`r|`n", "`r`n"
}

function Set-IbisTextBoxDisplayText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$TextBox,

        [AllowNull()]
        [string]$Text,

        [switch]$StripAnsi
    )

    $TextBox.Text = ConvertTo-IbisGuiDisplayText -Text $Text -StripAnsi:$StripAnsi
}

function Add-IbisTextBoxDisplayText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$TextBox,

        [AllowNull()]
        [string]$Text,

        [switch]$StripAnsi
    )

    $TextBox.AppendText((ConvertTo-IbisGuiDisplayText -Text $Text -StripAnsi:$StripAnsi))
}

function New-IbisSessionLogFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $logsDirectory = Join-Path $ProjectRoot 'logs'
    if (-not (Test-Path -LiteralPath $logsDirectory)) {
        New-Item -ItemType Directory -Path $logsDirectory -Force | Out-Null
    }

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH-mm-ssZ')
    $logPath = Join-Path $logsDirectory ('Ibis-{0}.log' -f $timestamp)
    New-Item -ItemType File -Path $logPath -Force | Out-Null
    $logPath
}

function Write-IbisGuiLogFileLine {
    [CmdletBinding()]
    param(
        [string]$LogFilePath,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($LogFilePath)) {
        return
    }

    try {
        $timestamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $line = '{0} [{1}] {2}' -f $timestamp, $Level.ToUpperInvariant(), $Message
        Add-Content -LiteralPath $LogFilePath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Get-IbisCommandLineHint {
    [CmdletBinding()]
    param(
        [object]$InputObject
    )

    $hints = New-Object System.Collections.Generic.List[string]
    $seen = @{}

    $visit = {
        param($Value)

        if ($null -eq $Value) {
            return
        }

        if ($Value -is [string]) {
            return
        }

        $valueType = $Value.GetType()
        if ($valueType.IsPrimitive -or $Value -is [decimal] -or $Value -is [datetime] -or $Value -is [guid] -or $Value -is [enum]) {
            return
        }

        if ($Value -is [System.Collections.IDictionary]) {
            foreach ($key in $Value.Keys) {
                if ([string]$key -eq 'CommandLine' -and -not [string]::IsNullOrWhiteSpace([string]$Value[$key])) {
                    $hint = [string]$Value[$key]
                    if (-not $seen.ContainsKey($hint)) {
                        $seen[$hint] = $true
                        $hints.Add($hint)
                    }
                }
                & $visit $Value[$key]
            }
            return
        }

        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
            foreach ($item in $Value) {
                & $visit $item
            }
            return
        }

        foreach ($property in @($Value.PSObject.Properties)) {
            if ($property.Name -eq 'CommandLine' -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                $hint = [string]$property.Value
                if (-not $seen.ContainsKey($hint)) {
                    $seen[$hint] = $true
                    $hints.Add($hint)
                }
            }
            & $visit $property.Value
        }
    }

    & $visit $InputObject
    $hints.ToArray()
}

function Add-IbisCommandLineHints {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$LogTextBox,

        [object]$Result,

        [string]$LogFilePath = $script:IbisCurrentLogFilePath
    )

    $sourceObjects = @()
    if ($null -ne $Result) {
        $sourceObjects += $Result
        if (-not [string]::IsNullOrWhiteSpace([string]$Result.JsonPath) -and (Test-Path -LiteralPath $Result.JsonPath -PathType Leaf)) {
            try {
                $sourceObjects += Get-Content -LiteralPath $Result.JsonPath -Raw | ConvertFrom-Json
            }
            catch {
                Add-IbisLogLine -LogTextBox $LogTextBox -LogFilePath $LogFilePath -Level 'WARN' -Message "Unable to read command line hints from summary JSON: $($Result.JsonPath)"
            }
        }
    }

    $seen = @{}
    foreach ($sourceObject in $sourceObjects) {
        foreach ($hint in @(Get-IbisCommandLineHint -InputObject $sourceObject)) {
            if ($seen.ContainsKey($hint)) {
                continue
            }
            $seen[$hint] = $true
            Add-IbisLogLine -LogTextBox $LogTextBox -LogFilePath $LogFilePath -Message "Command line hint: $hint"
        }
    }
}

function Get-IbisFileSystemSnapshot {
    [CmdletBinding()]
    param(
        [string]$RootPath,

        [int]$MaxItems = 50000
    )

    $items = @{}
    if ([string]::IsNullOrWhiteSpace($RootPath) -or -not (Test-Path -LiteralPath $RootPath -PathType Container)) {
        return [pscustomobject]@{
            RootPath = $RootPath
            Exists = $false
            Items = $items
            Truncated = $false
        }
    }

    $count = 0
    $truncated = $false
    foreach ($item in @(Get-ChildItem -LiteralPath $RootPath -Recurse -Force -ErrorAction SilentlyContinue)) {
        if ($count -ge $MaxItems) {
            $truncated = $true
            break
        }

        $itemType = 'Directory'
        $length = $null
        if (-not $item.PSIsContainer) {
            $itemType = 'File'
            $length = $item.Length
        }

        $items[$item.FullName] = [pscustomobject]@{
            Path = $item.FullName
            Type = $itemType
            Length = $length
            LastWriteTimeUtcTicks = $item.LastWriteTimeUtc.Ticks
        }
        $count++
    }

    [pscustomobject]@{
        RootPath = $RootPath
        Exists = $true
        Items = $items
        Truncated = $truncated
    }
}

function Add-IbisFileSystemChangeLog {
    [CmdletBinding()]
    param(
        [object]$Before,

        [object]$After,

        [string]$Context,

        [string]$LogFilePath = $script:IbisCurrentLogFilePath
    )

    if ($null -eq $Before -or $null -eq $After) {
        return
    }

    if ($Before.Truncated -or $After.Truncated) {
        Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Level 'WARN' -Message "File audit for ${Context}: snapshot was truncated; file operation log may be incomplete."
    }

    $beforeItems = $Before.Items
    $afterItems = $After.Items
    $created = New-Object System.Collections.Generic.List[object]
    $removed = New-Object System.Collections.Generic.List[object]
    $updated = New-Object System.Collections.Generic.List[object]

    foreach ($path in $afterItems.Keys) {
        if (-not $beforeItems.ContainsKey($path)) {
            $created.Add($afterItems[$path])
            continue
        }

        $beforeItem = $beforeItems[$path]
        $afterItem = $afterItems[$path]
        if ($afterItem.Type -eq 'File' -and (($beforeItem.Length -ne $afterItem.Length) -or ($beforeItem.LastWriteTimeUtcTicks -ne $afterItem.LastWriteTimeUtcTicks))) {
            $updated.Add($afterItem)
        }
    }

    foreach ($path in $beforeItems.Keys) {
        if (-not $afterItems.ContainsKey($path)) {
            $removed.Add($beforeItems[$path])
        }
    }

    Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message "File audit for ${Context}: $($created.Count) created, $($updated.Count) updated, $($removed.Count) removed."

    foreach ($item in @($created | Sort-Object Path)) {
        Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message "$($item.Type) created: $($item.Path)"
    }
    foreach ($item in @($updated | Sort-Object Path)) {
        Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message "$($item.Type) updated: $($item.Path)"
    }
    foreach ($item in @($removed | Sort-Object Path)) {
        Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message "$($item.Type) removed: $($item.Path)"
    }
}

function Get-IbisFileMoveHint {
    [CmdletBinding()]
    param(
        [object]$InputObject
    )

    $hints = New-Object System.Collections.Generic.List[object]
    $seen = @{}

    $visit = {
        param($Value)

        if ($null -eq $Value -or $Value -is [string]) {
            return
        }

        if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string]) -and -not ($Value -is [System.Collections.IDictionary])) {
            foreach ($item in $Value) {
                & $visit $item
            }
            return
        }

        $properties = @($Value.PSObject.Properties)
        $originalPath = ($properties | Where-Object { $_.Name -eq 'OriginalPath' } | Select-Object -First 1).Value
        $newPath = ($properties | Where-Object { $_.Name -eq 'NewPath' } | Select-Object -First 1).Value
        if (-not [string]::IsNullOrWhiteSpace([string]$originalPath) -and -not [string]::IsNullOrWhiteSpace([string]$newPath)) {
            $key = '{0}|{1}' -f $originalPath, $newPath
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $hints.Add([pscustomobject]@{
                    OriginalPath = [string]$originalPath
                    NewPath = [string]$newPath
                })
            }
        }

        foreach ($property in $properties) {
            & $visit $property.Value
        }
    }

    & $visit $InputObject
    $hints.ToArray()
}

function Add-IbisFileOperationHints {
    [CmdletBinding()]
    param(
        [object]$Result,

        [string]$LogFilePath = $script:IbisCurrentLogFilePath
    )

    $sourceObjects = @()
    if ($null -ne $Result) {
        $sourceObjects += $Result
        if (-not [string]::IsNullOrWhiteSpace([string]$Result.JsonPath) -and (Test-Path -LiteralPath $Result.JsonPath -PathType Leaf)) {
            try {
                $sourceObjects += Get-Content -LiteralPath $Result.JsonPath -Raw | ConvertFrom-Json
            }
            catch {
                Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Level 'WARN' -Message "Unable to read file operation hints from summary JSON: $($Result.JsonPath)"
            }
        }
    }

    $seen = @{}
    foreach ($sourceObject in $sourceObjects) {
        foreach ($hint in @(Get-IbisFileMoveHint -InputObject $sourceObject)) {
            $key = '{0}|{1}' -f $hint.OriginalPath, $hint.NewPath
            if ($seen.ContainsKey($key)) {
                continue
            }
            $seen[$key] = $true
            Write-IbisGuiLogFileLine -LogFilePath $LogFilePath -Message "File moved/renamed: $($hint.OriginalPath) -> $($hint.NewPath)"
        }
    }
}

function New-IbisWindowIcon {
    $iconBase64 = @(
        'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFiUA'
        'ABYlAUlSJPAAAA2PSURBVHhe7ZsJeBRVtsf/1Z1Oujsb2SALIQsJEkyQdUYkBBIgDpBgAAcdxIAiMMAICiIO4BBEHTQSIOwIGp8g'
        'jCigw74LCaAiawjZO5B9hWy9d985VZaPzwlolg7S7/n7vqRv3aquqvu/555z7q1q/M7v/M7/azjxs1VsWr+apV+5BKPRBG/fLngz'
        '4Z02ne+3oNU3PGp4JGtsUMPVzRMSTorGhlrU3C7HzDmvYNILL1mNEK260aefGsV0ajO6Bj0Ko9kEs8kEiQRQqxuRfu08vrty3WoE'
        'oNtuGQf3f8UqyysQEPiI0OP5+Vfpsxg6nQ5KhSPcXDzx9pKFTDz8oadFAuzetZOtWZmIzp0DYDCYUVKSj6ry2ygqVMFkNMJE1qC0'
        'd0RRUZH4jYefZguwaP48lvj2u3B194ezixf0Zh0U1OOcxAi5nRwcx4Hx/c7RKalsLTRLgPgJ49jp02no//gQaqw9qmtKwVFjO3Xq'
        'guDgR9GlSwi1W0oNZzAYDSSMQvzmw8+vChD/3FhWXFSBPn2eQHl5EbIyL8JM5i63lUMms4PCvgMkUhvhWAn1vkGvhauri7BtDfyi'
        'AAtfn8sKcosRFtYfRbdyUFZajKDAECjJCrIyLyNflUERwCAcy8j++WGg1+vg6eUt1FkD9xVgz5efsxNHjqFXrz9Sw2+Rp6+Cm6sn'
        'Cm7mwizRIu7Po9G3Xyg5wBzY2MiExpvJCZqMeoT0CBXP8vBzXwG2bFqP4KBQaLQ6lFUU0tjWADItFiz+Oz7duYubNXseFzk0mvZr'
        'KQeQCAJotWoa/7YIj4i07jxgw9pk1lCngYubB8qp8RUlJeQD+mH314e5P42M/d/GXfj+W9iRL+DNXyLhUFNTiS6BAeJe6+CeApw6'
        'dhRd/IMp1JlQVqxC9IgRSN64pUmvqnLz4OHuRY2XQi5XQKetR/igweJe66CJAGe+OclKS8rRwaUjCgty4B8QgFXrNt/TpHOzc+Hg'
        '6Ay9QQd1YwOMJj2en/yi9SQBRBMBLv5AZk29qaexXXenDJ9/deCeDXp32RImsVFASg5QQuO/pqaKxPIT91oPTQTIysqmGV4n6t0r'
        '+POEiWJtU86cOgVvbz+aChspD5CgtrYa/f4wQNxrPTQRoKaqBlq1Gi6uDpj96vx79v5HWzYydaMODg7OggC8E+TN/5GQHuIR1kMT'
        'AXjT5z3/xElTxJqmnE1Ng5ubJ0w0DZbL5dDrtJArbBAZNcyqxj9PEwGqaSx7eXsi7unx921MGYXFjp28IJXZoLAwGwUFGWisbxD3'
        'WhdNBaioxKjRT4lbTTl18hhraKhHIWWEF84dh8LBFsfPnEXMhOfwl7FxVrMOcF8Wvjb3Fxux7ZOPWL8ewWzZe8vY2i3rWdr1q+QC'
        'fuSDTRvZuDHj2MrEVVYjRIvHbOI7b7Gy2tv44L0ksebnnLxyDcf37IUqOxsd3VwR0i0A02e/+tD6hhbf2Mrl7zLOxRlzps0Ua+6S'
        'RiEUsEF4967cicvpLDfzGm5lpSMv4wZkNFUOjxiMaTNftjpH+TPWJH3A3l+9UjT6e3M2O1cs3eXrE0fZnNdeZmNiolniu28/NEOk'
        'iRP8NeQKOTSaux7/ky/+RT1/Xdz6kQHBXcXSXWIjh2FVYjJeeT8RGimHmOgotnXTht9ciBYL4O3tTclShbgF7P14Kxpqa8WtXyci'
        'pCfefH0h5q9IxLmL5/H8M2N/UxFaLMCop8Zy6tp6cQtQ2tvDzs5O3Go+EaF98OHGFITHxOLJ4UPY5zs+/U2EaLEAPHqNViwBBpNB'
        'WA5vLdMmTsbi1aux7bNP8c+3Fj9wEVolAL/4eeT780LZTqYAWNscezgNi6++PozsW4WY/+rsBypCqwRwdXFBduZPjo+DUa8Xy23j'
        'ow9TYLaRYsZLLzwwEVolwKM068u4eFEo6yVmlJeWCmVL8MH7SbCxV+Ll6VMeiAitEuDFGbO4ypuFQtk3IBCXL3wnlC1F8qq10IJh'
        'ycIF7S5CqwTgsbWxwc6TBzFo5Chknr+A4xmXcOJGOs5kZYhHtI3NG7fiKp1vO809xKp2odUCDBwUgdR9BxEXPoTr3r8Ptq1Mhplf'
        'GKGIcCz9ClLz+LS4bezevQ8pn2wVt9oHqfjZYg4cPLzUs4NzwrrdXyb0GjSQJkC7YVJr4NOtmzDB4ChS3KyuQn55CW7drkZZfR2K'
        'au6gMznQlsAclFDo9QnpmdlLxSqL0moL4BkaGYXt69bAbJZgRsJSpB45ggvHjlJU5GAwGqHXazH00cc4G5kd7tTWwchMOJOThZPX'
        'r+G8Klc4x9ncLOHzfsSPexYuXXyRnJTYLkOh1RbAc+jo8aWP+Pok6AxaBPfsjccG9MOeDzejpvo2gsNCYdBpYK69k+BPVjH68YHc'
        'hOkvJfCP0G1tbaE16JBfUU6WIkUBpdaqqjKU1tWikKzlZk0Vuri6i1cBnPw64+P165Gbp7K4FVhkavrkkAgW+XQcwkfFQd1Yi4+X'
        'vwdGvkAhl6OMkhuJnQ0Gj4zF67PmcKdzc5jZaCB/YSI/T2LIbGEy8daip5RaLpxPR2VbG1tIaNLUqG7EiMf6cm8spSzxTh2Wr0y2'
        '6HS6TUPgJw6fOs19e+gIPkteCUcHR+zYtpMLHzEK7l4+WLA2GdPfXIxThw8hPv45VldZjiHde3AysgIZNVKjISsxA9FhvTn++aLR'
        'YCDhFDSE9NBoNXCgucaJjHQ2fPx4ZBcV4NyZ0xYdChZVc+7MGSwrLwdjpk1F9z79hTq9Rk0NlcHGVor9/7MNZ08eQ8SfRiBi9GhE'
        'hYRxafl5jH+xIvPSD5BTY/2690C4X2CT+0rNzWarF7+B6kIVTqRdsth9W1QAnkFP9GUuLu7o4O6BaOo1/x7dER7QjUvLy2ESmQTF'
        'qnyc+HI38jNvoA+F0j9GDadvmZE0bx4cnZ2hcHFCJ28fePsHwsnZBRq1GlWlRbh8Ng29BoRD7qjEzYuXkbLzC4vcu8UFeH7iMyw6'
        'fgJU1zNxbv9BeHT2Qe8hg9F/8DDxiB8pycvDmQP/xi0SRNvQQPuHYGhcHLLSr6PkpgoVhYXQqrWwo6Hi4OGC0H79ydH2JTEDudkL'
        '5jJJgxqr1m1s8/1bXIBl/1jEGqinYye+gKxr32NP8nq4enigtKQYgaFh6BwUjK7dQuAZ4A9bhVL8VvMwaLWQkWPlSUn8J5wkNkha'
        'u6FNbbC4ADwxMcPYnMQV+P70N1CrCvH28kQuZuRQFhE7Ghk/XEJ1URFdmKKEkxONewcoHJ0hkUnh4uYKpRMNA6UDvP384ePbRYgg'
        'Bp2J3ASDnhwj+UnIFQ6CJUz+6wusAyfDqg33fnrdHNpFgJlTJrGc/Bw01N7BsxMmYc5rC7ixcTFs4oLX0dGzMz5Leg8jo6IRM2Yc'
        'd+zwQaYmc66rrUFlZRXqa+spaapG5e0qaHR6ePkHIPKpOHj5BUBLEYN/G0VKoZOHF2HGnFlMT3nH1m2ftaot7SLAasraDuw9BKXC'
        'Ho+E+mP5imTuo00b2N59XyEoNASqjCzsuc9j9/8m5cNNbO/uLxA+ZiweH/4kjDothU3zz4bPvu0pyPvuB+z6en+L22ORPOC/UdpT'
        'HDfpKPTJUVxYLNS9OH0GV0WOLSP1Eo1lnVDXHCZPnc7tPXiUO/H5DuRevUA+wI4ySTsaFlrhpQyemOcmo9fQKIyIjGhxjtAuAnh4'
        'dBQem/Mpb0P93QXUDu4dERgcBmZqeS4z/43F+GLjZuG1PP7cfNIkI7+h16qF/YvmzONip05B1KAnWnTydhHAmeI3b6ZSiVR4c+wn'
        '+DfK+eyO0eSppUQOi+YeC+2JXZs3CguxfKrN/06B/+NJVeWzngMGIXbSJMRERzZbhHYRgPfUUjozR2HKoL+7YiyzldF/JqyhHjqw'
        'r8Vm8E5iEqe6nI7CjAwY+fyZUCrkZAWNQpnn1Zemc2HhA/FOQvNWmNtFgCFRw/moRSJIYBJ7iEdCgvB3RRIIL1a1hn8fPsZtX78W'
        'BnUDhUMFRQodbOX2wr7cq5fxCiVJd27XQFWgEup+jXYRYNsnW5mZGimRSqEzGMXaH1+nJduFlDcRsQdbw8wZf8PqRYuFqXRxbo5Y'
        'C7h6dsTV1HPwc/fDlpTtzYoILQ4bv8SyJYvYNeqFstJKeHbyg73CGaVlBfD0cUHK9l3cs+NimdzWHRUVhZg0JR7PTIhv9fU3rVvD'
        '9h89gA40f8i5egVBvXqja0gPZF26gR07mtd4HotYwJJFb7CB/fqytG++pYzNHkGBPaGwc6JONsLb2x/lJXV4om8YqywrJbO1h4Gm'
        'vHyEaAvTZ73MzZw6C8F+QTh36QY3NX4KjFUNGD/+L+IRzaPNFrB+zSq2esUKhHTvDbWmATpdo/DDCaVSKSx0mMjUOf53BDQUPDw6'
        'kWUEoKS0AG4eSjLTnRa1wNbQ5htISz3NkpMS4ermhu49QuDr60eNd8DHWzagpLgMXp6+sFfy+b0TOPIJjASxo2igys9Eo7oWcxcs'
        'wIhRd98/ftC0y4W/OXmczZ4xDcFBPaGm+bxer6F4TV5fwgnPFW1tFXBwcBJWg2rry5H6reUWOFpKu1149sxpTCJl6No1GD4+vtRg'
        'RyEENtLcv7KyAjcL8pGfm4cQSm6WvGV9P7j8nf8bAP8B9w2d8s4x6N0AAAAASUVORK5CYII='
    ) -join ''

    $bytes = [Convert]::FromBase64String($iconBase64)
    $stream = New-Object System.IO.MemoryStream(,$bytes)
    $bitmap = $null
    $handle = [IntPtr]::Zero
    try {
        $bitmap = [System.Drawing.Bitmap]::FromStream($stream)
        $handle = $bitmap.GetHicon()
        $icon = [System.Drawing.Icon]::FromHandle($handle)
        $icon.Clone()
    }
    finally {
        if ($handle -ne [IntPtr]::Zero) {
            if (-not ('IbisNativeMethods' -as [type])) {
                Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class IbisNativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
'@
            }
            [void][IbisNativeMethods]::DestroyIcon($handle)
        }
        if ($bitmap) {
            $bitmap.Dispose()
        }
        $stream.Dispose()
    }
}

function Read-IbisProgressEvents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgressPath,

        [int]$SkipCount = 0
    )

    if ([string]::IsNullOrWhiteSpace($ProgressPath) -or -not (Test-Path -LiteralPath $ProgressPath)) {
        return [pscustomobject]@{
            Events = @()
            LineCount = $SkipCount
        }
    }

    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.FileStream]::new(
            $ProgressPath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
        $lines = @($reader.ReadToEnd() -split "\r?\n")
    }
    catch {
        return [pscustomobject]@{
            Events = @()
            LineCount = $SkipCount
        }
    }
    finally {
        if ($reader) {
            $reader.Dispose()
        }
        elseif ($stream) {
            $stream.Dispose()
        }
    }

    $lines = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($lines.Count -le $SkipCount) {
        return [pscustomobject]@{
            Events = @()
            LineCount = $lines.Count
        }
    }

    $events = @()
    for ($i = $SkipCount; $i -lt $lines.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace($lines[$i])) {
            continue
        }

        try {
            $events += ($lines[$i] | ConvertFrom-Json)
        }
        catch {
            $events += [pscustomobject]@{
                Time = (Get-Date).ToString('s')
                ToolId = ''
                ToolName = ''
                Stage = 'Progress'
                Message = $lines[$i]
                Index = 0
                Total = 0
                Status = 'Info'
            }
        }
    }

    [pscustomobject]@{
        Events = $events
        LineCount = $lines.Count
    }
}

function Set-IbisProcessingControlState {
    param(
        [string]$ControlPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Running', 'Paused', 'CancelRequested')]
        [string]$State
    )

    if ([string]::IsNullOrWhiteSpace($ControlPath)) {
        return
    }

    $directory = Split-Path -Path $ControlPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $payload = [pscustomobject]@{
        State = $State
        Updated = (Get-Date).ToString('s')
    } | ConvertTo-Json -Compress

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.FileStream]::new(
                $ControlPath,
                [System.IO.FileMode]::Create,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
            $writer.Write($payload)
            return
        }
        catch {
            if ($attempt -eq 5) {
                throw
            }
            Start-Sleep -Milliseconds (25 * $attempt)
        }
        finally {
            if ($writer) {
                $writer.Dispose()
            }
            elseif ($stream) {
                $stream.Dispose()
            }
        }
    }
}

function Start-IbisToolInstallRunspace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [string]$ProgressPath
    )

    $coreModulePath = Join-Path $ProjectRoot 'modules\Ibis.Core.psm1'
    $scriptBlock = {
        param(
            [string]$ProjectRoot,
            [string]$ToolsRoot,
            [string]$CoreModulePath,
            [string]$ProgressPath
        )

        $ErrorActionPreference = 'Stop'
        Import-Module $CoreModulePath -Force
        $config = Get-IbisConfig -ProjectRoot $ProjectRoot
        $toolDefinitions = @(Get-IbisToolDefinition -ProjectRoot $ProjectRoot -Config $config)
        @(Invoke-IbisInstallMissingTools -ToolsRoot $ToolsRoot -ToolDefinitions $toolDefinitions -ProgressPath $ProgressPath)
    }

    $powershell = [PowerShell]::Create()
    [void]$powershell.AddScript($scriptBlock)
    [void]$powershell.AddArgument($ProjectRoot)
    [void]$powershell.AddArgument($ToolsRoot)
    [void]$powershell.AddArgument($coreModulePath)
    [void]$powershell.AddArgument($ProgressPath)

    [pscustomobject]@{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Started = Get-Date
    }
}

function Start-IbisHayabusaRulesUpdateRunspace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot
    )

    $coreModulePath = Join-Path $ProjectRoot 'modules\Ibis.Core.psm1'
    $scriptBlock = {
        param(
            [string]$ProjectRoot,
            [string]$ToolsRoot,
            [string]$CoreModulePath
        )

        $ErrorActionPreference = 'Stop'
        Import-Module $CoreModulePath -Force
        $config = Get-IbisConfig -ProjectRoot $ProjectRoot
        $toolDefinitions = @(Get-IbisToolDefinition -ProjectRoot $ProjectRoot -Config $config)
        Invoke-IbisHayabusaRuleUpdate -ToolsRoot $ToolsRoot -ToolDefinitions $toolDefinitions
    }

    $powershell = [PowerShell]::Create()
    [void]$powershell.AddScript($scriptBlock)
    [void]$powershell.AddArgument($ProjectRoot)
    [void]$powershell.AddArgument($ToolsRoot)
    [void]$powershell.AddArgument($coreModulePath)

    [pscustomobject]@{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Started = Get-Date
    }
}

function Start-IbisProcessingRunspace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [object[]]$Modules,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [string]$ProgressPath,

        [string]$ControlPath
    )

    $coreModulePath = Join-Path $ProjectRoot 'modules\Ibis.Core.psm1'
    $scriptBlock = {
        param(
            [string]$ProjectRoot,
            [string]$ToolsRoot,
            [object[]]$ToolDefinitions,
            [object[]]$Modules,
            [string]$SourceRoot,
            [string]$OutputRoot,
            [string]$Hostname,
            [string]$CoreModulePath,
            [string]$ProgressPath,
            [string]$ControlPath
        )

        $ErrorActionPreference = 'Stop'
        Import-Module $CoreModulePath -Force

        function Get-IbisProcessingSnapshot {
            param([string]$RootPath)

            $items = @{}
            if ([string]::IsNullOrWhiteSpace($RootPath) -or -not (Test-Path -LiteralPath $RootPath -PathType Container)) {
                return [pscustomobject]@{ RootPath = $RootPath; Exists = $false; Items = $items; Truncated = $false }
            }

            $count = 0
            $truncated = $false
            foreach ($item in @(Get-ChildItem -LiteralPath $RootPath -Recurse -Force -ErrorAction SilentlyContinue)) {
                if ($count -ge 50000) {
                    $truncated = $true
                    break
                }

                $itemType = 'Directory'
                $length = $null
                if (-not $item.PSIsContainer) {
                    $itemType = 'File'
                    $length = $item.Length
                }

                $items[$item.FullName] = [pscustomobject]@{
                    Path = $item.FullName
                    Type = $itemType
                    Length = $length
                    LastWriteTimeUtcTicks = $item.LastWriteTimeUtc.Ticks
                }
                $count++
            }

            [pscustomobject]@{ RootPath = $RootPath; Exists = $true; Items = $items; Truncated = $truncated }
        }

        function Write-IbisProcessingFileAudit {
            param(
                [object]$Before,
                [object]$After,
                [string]$Context,
                [int]$Index,
                [int]$Total
            )

            if ($null -eq $Before -or $null -eq $After) {
                return
            }

            if ($Before.Truncated -or $After.Truncated) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'file-audit' -ToolName $Context -Stage 'File audit' -Message 'Snapshot was truncated; file operation log may be incomplete.' -Index $Index -Total $Total -Status 'Audit'
            }

            $created = New-Object System.Collections.Generic.List[object]
            $updated = New-Object System.Collections.Generic.List[object]
            $removed = New-Object System.Collections.Generic.List[object]

            foreach ($path in $After.Items.Keys) {
                if (-not $Before.Items.ContainsKey($path)) {
                    $created.Add($After.Items[$path])
                    continue
                }

                $beforeItem = $Before.Items[$path]
                $afterItem = $After.Items[$path]
                if ($afterItem.Type -eq 'File' -and (($beforeItem.Length -ne $afterItem.Length) -or ($beforeItem.LastWriteTimeUtcTicks -ne $afterItem.LastWriteTimeUtcTicks))) {
                    $updated.Add($afterItem)
                }
            }

            foreach ($path in $Before.Items.Keys) {
                if (-not $After.Items.ContainsKey($path)) {
                    $removed.Add($Before.Items[$path])
                }
            }

            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'file-audit' -ToolName $Context -Stage 'File audit' -Message "$($created.Count) created, $($updated.Count) updated, $($removed.Count) removed." -Index $Index -Total $Total -Status 'Audit'
            foreach ($item in @($created | Sort-Object Path)) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'file-audit' -ToolName $Context -Stage 'File audit' -Message "$($item.Type) created: $($item.Path)" -Index $Index -Total $Total -Status 'Audit'
            }
            foreach ($item in @($updated | Sort-Object Path)) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'file-audit' -ToolName $Context -Stage 'File audit' -Message "$($item.Type) updated: $($item.Path)" -Index $Index -Total $Total -Status 'Audit'
            }
            foreach ($item in @($removed | Sort-Object Path)) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'file-audit' -ToolName $Context -Stage 'File audit' -Message "$($item.Type) removed: $($item.Path)" -Index $Index -Total $Total -Status 'Audit'
            }
        }

        function Get-IbisProcessingControlState {
            if ([string]::IsNullOrWhiteSpace($ControlPath) -or -not (Test-Path -LiteralPath $ControlPath -PathType Leaf)) {
                return 'Running'
            }

            $stream = $null
            $reader = $null
            try {
                $stream = [System.IO.FileStream]::new(
                    $ControlPath,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite
                )
                $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
                $raw = $reader.ReadToEnd()
                if ([string]::IsNullOrWhiteSpace($raw)) {
                    return 'Running'
                }

                $control = $raw | ConvertFrom-Json
                if ([string]::IsNullOrWhiteSpace([string]$control.State)) {
                    return 'Running'
                }

                [string]$control.State
            }
            catch {
                'Running'
            }
            finally {
                if ($reader) {
                    $reader.Dispose()
                }
                elseif ($stream) {
                    $stream.Dispose()
                }
            }
        }

        function Wait-IbisProcessingControl {
            param(
                [string]$NextModuleName,
                [int]$Index,
                [int]$Total
            )

            $pauseLogged = $false
            while ($true) {
                $state = Get-IbisProcessingControlState
                if ($state -eq 'CancelRequested') {
                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'ibis' -ToolName 'Ibis' -Stage 'Cancelled' -Message "Processing cancelled before starting $NextModuleName." -Index $Index -Total $Total -Status 'Cancelled'
                    return 'CancelRequested'
                }
                if ($state -ne 'Paused') {
                    if ($pauseLogged) {
                        Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'ibis' -ToolName 'Ibis' -Stage 'Resumed' -Message "Processing resumed before $NextModuleName." -Index $Index -Total $Total -Status 'Running'
                    }
                    return 'Running'
                }

                if (-not $pauseLogged) {
                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'ibis' -ToolName 'Ibis' -Stage 'Paused' -Message "Processing paused before $NextModuleName." -Index $Index -Total $Total -Status 'Paused'
                    $pauseLogged = $true
                }
                Start-Sleep -Milliseconds 750
            }
        }

        function Get-IbisProcessingCommandLineHint {
            param([object]$InputObject)

            $hints = New-Object System.Collections.Generic.List[string]
            $seen = @{}

            $visit = {
                param($Value)

                if ($null -eq $Value -or $Value -is [string]) {
                    return
                }

                $valueType = $Value.GetType()
                if ($valueType.IsPrimitive -or $Value -is [decimal] -or $Value -is [datetime] -or $Value -is [guid] -or $Value -is [enum]) {
                    return
                }

                if ($Value -is [System.Collections.IDictionary]) {
                    foreach ($key in $Value.Keys) {
                        if ([string]$key -eq 'CommandLine' -and -not [string]::IsNullOrWhiteSpace([string]$Value[$key])) {
                            $hint = [string]$Value[$key]
                            if (-not $seen.ContainsKey($hint)) {
                                $seen[$hint] = $true
                                $hints.Add($hint)
                            }
                        }
                        & $visit $Value[$key]
                    }
                    return
                }

                if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
                    foreach ($item in $Value) {
                        & $visit $item
                    }
                    return
                }

                foreach ($property in @($Value.PSObject.Properties)) {
                    if ($property.Name -eq 'CommandLine' -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $hint = [string]$property.Value
                        if (-not $seen.ContainsKey($hint)) {
                            $seen[$hint] = $true
                            $hints.Add($hint)
                        }
                    }
                    & $visit $property.Value
                }
            }

            & $visit $InputObject
            $hints.ToArray()
        }

        function Write-IbisProcessingCommandLineHints {
            param(
                [object]$Result,
                [string]$ModuleId,
                [string]$ModuleName,
                [int]$Index,
                [int]$Total
            )

            if ($null -eq $Result) {
                return
            }

            $sourceObjects = @($Result)
            if (-not [string]::IsNullOrWhiteSpace([string]$Result.JsonPath) -and (Test-Path -LiteralPath $Result.JsonPath -PathType Leaf)) {
                try {
                    $sourceObjects += Get-Content -LiteralPath $Result.JsonPath -Raw | ConvertFrom-Json
                }
                catch {
                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ModuleId -ToolName $ModuleName -Stage 'Command line hint' -Message "Unable to read command line hints from summary JSON: $($Result.JsonPath)" -Index $Index -Total $Total -Status 'Warning'
                }
            }

            $seen = @{}
            foreach ($sourceObject in $sourceObjects) {
                foreach ($hint in @(Get-IbisProcessingCommandLineHint -InputObject $sourceObject)) {
                    if ($seen.ContainsKey($hint)) {
                        continue
                    }
                    $seen[$hint] = $true
                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ModuleId -ToolName $ModuleName -Stage 'Command line hint' -Message $hint -Index $Index -Total $Total -Status 'Info'
                }
            }
        }

        $currentOutputRoot = $OutputRoot
        $currentHostname = $Hostname
        $preserveBlankHostname = [string]::IsNullOrWhiteSpace($Hostname)
        $moduleResults = @()
        $total = $Modules.Count

        for ($moduleIndex = 0; $moduleIndex -lt $Modules.Count; $moduleIndex++) {
            $module = $Modules[$moduleIndex]
            $index = $moduleIndex + 1
            $moduleName = $module.name
            $moduleId = $module.id

            $controlState = Wait-IbisProcessingControl -NextModuleName $moduleName -Index $index -Total $total
            if ($controlState -eq 'CancelRequested') {
                break
            }

            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Started' -Message "Running $moduleName ($index of $total)" -Index $index -Total $total -Status 'Started'
            $snapshotBefore = Get-IbisProcessingSnapshot -RootPath $currentOutputRoot
            $result = $null
            $errorMessage = $null

            try {
                switch ($moduleId) {
                    'system-summary' { $result = Invoke-IbisSystemSummary -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'velociraptor-results' { $result = Invoke-IbisVelociraptorResultsCopy -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'registry' { $result = Invoke-IbisWindowsRegistryHives -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'amcache' { $result = Invoke-IbisAmcache -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'appcompatcache' { $result = Invoke-IbisAppCompatCache -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'prefetch' { $result = Invoke-IbisPrefetch -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'ntfs-metadata' { $result = Invoke-IbisNtfsMetadata -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'srum' { $result = Invoke-IbisSrum -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'user-artifacts' { $result = Invoke-IbisUserArtifacts -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'eventlogs' { $result = Invoke-IbisEvtxECmdEventLogs -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'duckdb-eventlogs' { $result = Invoke-IbisDuckDbEventLogSummary -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -OutputRoot $currentOutputRoot -Hostname $currentHostname -ProjectRoot $ProjectRoot }
                    'hayabusa' { $result = Invoke-IbisHayabusaEventLogs -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'takajo' { $result = Invoke-IbisTakajoEventLogs -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'chainsaw' { $result = Invoke-IbisChainsawEventLogs -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'ual' { $result = Invoke-IbisUserAccessLogsSum -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'browser-history' { $result = Invoke-IbisBrowsingHistoryView -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'forensic-webhistory' { $result = Invoke-IbisForensicWebHistory -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    'usb' { $result = Invoke-IbisParseUsbArtifacts -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -SourceRoot $SourceRoot -OutputRoot $currentOutputRoot -Hostname $currentHostname }
                    default { throw "Processing module is not implemented: $moduleId" }
                }

                if ($null -ne $result) {
                    if (-not $preserveBlankHostname -and $result.HostName -and $result.HostName -ne 'Unknown') {
                        $currentHostname = $result.HostName
                    }
                    if (-not $preserveBlankHostname -and $result.HostOutputRoot) {
                        $currentOutputRoot = $result.HostOutputRoot
                    }

                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Completed' -Message "$($result.ModuleId): $($result.Status) - $($result.Message)" -Index $index -Total $total -Status $result.Status
                    Write-IbisProcessingCommandLineHints -Result $result -ModuleId $moduleId -ModuleName $moduleName -Index $index -Total $total
                    if ($result.OutputPath) { Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Output' -Message $result.OutputPath -Index $index -Total $total -Status 'Info' }
                    if ($result.OutputDirectory) { Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Output' -Message $result.OutputDirectory -Index $index -Total $total -Status 'Info' }
                    if ($result.JsonPath) { Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Summary' -Message $result.JsonPath -Index $index -Total $total -Status 'Info' }
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $moduleId -ToolName $moduleName -Stage 'Failed' -Message "$moduleId failed: $errorMessage" -Index $index -Total $total -Status 'Failed'
            }

            $snapshotAfter = Get-IbisProcessingSnapshot -RootPath $snapshotBefore.RootPath
            Write-IbisProcessingFileAudit -Before $snapshotBefore -After $snapshotAfter -Context $moduleName -Index $index -Total $total
            $moduleResults += [pscustomobject]@{
                ModuleId = $moduleId
                ModuleName = $moduleName
                Result = $result
                ErrorMessage = $errorMessage
                    Hostname = $currentHostname
                OutputRoot = $currentOutputRoot
            }
        }

        if ((Get-IbisProcessingControlState) -eq 'CancelRequested') {
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'ibis' -ToolName 'Ibis' -Stage 'Cancelled' -Message 'Processing module run cancelled.' -Index $total -Total $total -Status 'Cancelled'
        }
        else {
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'ibis' -ToolName 'Ibis' -Stage 'Completed' -Message 'Processing module run complete.' -Index $total -Total $total -Status 'Completed'
        }
        $moduleResults
    }

    $powershell = [PowerShell]::Create()
    [void]$powershell.AddScript($scriptBlock)
    [void]$powershell.AddArgument($ProjectRoot)
    [void]$powershell.AddArgument($ToolsRoot)
    [void]$powershell.AddArgument($ToolDefinitions)
    [void]$powershell.AddArgument($Modules)
    [void]$powershell.AddArgument($SourceRoot)
    [void]$powershell.AddArgument($OutputRoot)
    [void]$powershell.AddArgument($Hostname)
    [void]$powershell.AddArgument($coreModulePath)
    [void]$powershell.AddArgument($ProgressPath)
    [void]$powershell.AddArgument($ControlPath)

    [pscustomobject]@{
        PowerShell = $powershell
        Handle = $powershell.BeginInvoke()
        Started = Get-Date
    }
}

function Stop-IbisToolInstallRunspace {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Operation
    )

    if ($Operation.PowerShell) {
        $Operation.PowerShell.Dispose()
    }
}

function Show-IbisGui {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $config = Get-IbisConfig -ProjectRoot $ProjectRoot
    $toolDefinitions = @(Get-IbisToolDefinition -ProjectRoot $ProjectRoot -Config $config)
    $logsDirectory = Join-Path $ProjectRoot 'logs'
    $sessionLogPath = New-IbisSessionLogFile -ProjectRoot $ProjectRoot
    $script:IbisCurrentLogFilePath = $sessionLogPath
    Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message 'Ibis GUI session started.'
    Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Project root: $ProjectRoot"
    Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Session log: $sessionLogPath"
    $completionBeepEnabled = $true
    if ($null -ne $config.PSObject.Properties['completionBeepEnabled']) {
        $completionBeepEnabled = [bool]$config.completionBeepEnabled
    }
    $appVersion = '0.0.0'
    if ($null -ne $config.PSObject.Properties['version'] -and -not [string]::IsNullOrWhiteSpace([string]$config.version)) {
        $appVersion = [string]$config.version
    }
    $changelogPath = Join-Path $ProjectRoot 'CHANGELOG.md'

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Ibis v${appVersion}: Tool Runner"
    $form.Size = New-Object System.Drawing.Size(820, 820)
    $form.MinimumSize = New-Object System.Drawing.Size(820, 760)
    $form.StartPosition = 'CenterScreen'
    try {
        $form.Icon = New-IbisWindowIcon
    }
    catch {
        Write-Warning "Unable to load embedded Ibis icon: $($_.Exception.Message)"
    }

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $form.Controls.Add($tabs)

    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready'
    [void]$statusStrip.Items.Add($statusLabel)
    $form.Controls.Add($statusStrip)

    $infoTab = New-Object System.Windows.Forms.TabPage
    $infoTab.Text = 'Info'
    $tabs.TabPages.Add($infoTab)

    $infoTitleLabel = New-Object System.Windows.Forms.Label
    $infoTitleLabel.Text = 'Ibis'
    $infoTitleLabel.Font = New-Object System.Drawing.Font($infoTitleLabel.Font.FontFamily, 18, [System.Drawing.FontStyle]::Bold)
    $infoTitleLabel.Location = New-Object System.Drawing.Point(18, 18)
    $infoTitleLabel.Size = New-Object System.Drawing.Size(620, 34)
    $infoTab.Controls.Add($infoTitleLabel)

    $infoTaglineLabel = New-Object System.Windows.Forms.Label
    $infoTaglineLabel.Text = 'A bin chicken for Windows forensic artefacts.'
    $infoTaglineLabel.Font = New-Object System.Drawing.Font($infoTaglineLabel.Font.FontFamily, 10, [System.Drawing.FontStyle]::Italic)
    $infoTaglineLabel.Location = New-Object System.Drawing.Point(20, 54)
    $infoTaglineLabel.Size = New-Object System.Drawing.Size(620, 24)
    $infoTab.Controls.Add($infoTaglineLabel)

    $infoLogoPictureBox = New-Object System.Windows.Forms.PictureBox
    $infoLogoPictureBox.Location = New-Object System.Drawing.Point(664, 16)
    $infoLogoPictureBox.Size = New-Object System.Drawing.Size(78, 78)
    $infoLogoPictureBox.SizeMode = 'Zoom'
    try {
        $infoLogoIcon = New-IbisWindowIcon
        $infoLogoPictureBox.Image = $infoLogoIcon.ToBitmap()
    }
    catch {
    }
    $infoTab.Controls.Add($infoLogoPictureBox)

    $infoTextBox = New-Object System.Windows.Forms.TextBox
    $infoTextBox.Location = New-Object System.Drawing.Point(18, 112)
    $infoTextBox.Size = New-Object System.Drawing.Size(744, 126)
    $infoTextBox.Multiline = $true
    $infoTextBox.ReadOnly = $true
    $infoTextBox.BorderStyle = 'FixedSingle'
    $infoTextBox.BackColor = [System.Drawing.SystemColors]::Window
    $infoTextBox.Text = "Ibis can download and prepare common digital forensic tools, then run selected processing modules against a Windows evidence source.`r`n`r`nSelect a source directory that represents a Windows system, such as a mounted disk image, a Velociraptor collection, or a KAPE triage collection. Then choose an output folder and select the processing modules you want to run.`r`n`r`nIbis records skipped artefacts, tool failures, and output locations so an analyst can review or manually rerun steps where needed."
    $infoTab.Controls.Add($infoTextBox)

    $disclaimerGroup = New-Object System.Windows.Forms.GroupBox
    $disclaimerGroup.Text = 'Disclaimer and licence'
    $disclaimerGroup.Location = New-Object System.Drawing.Point(18, 258)
    $disclaimerGroup.Size = New-Object System.Drawing.Size(744, 220)
    $infoTab.Controls.Add($disclaimerGroup)

    $disclaimerTextBox = New-Object System.Windows.Forms.TextBox
    $disclaimerTextBox.Location = New-Object System.Drawing.Point(14, 28)
    $disclaimerTextBox.Size = New-Object System.Drawing.Size(716, 176)
    $disclaimerTextBox.Multiline = $true
    $disclaimerTextBox.ReadOnly = $true
    $disclaimerTextBox.BorderStyle = 'None'
    $disclaimerTextBox.BackColor = [System.Drawing.SystemColors]::Control
    $disclaimerTextBox.Text = "Ibis is provided for digital forensic triage and analysis support. You are responsible for deciding whether it is appropriate for your evidence, environment, and legal or operational requirements.`r`n`r`nNo responsibility is accepted for data loss, evidence alteration, missed artefacts, incorrect results, tool behaviour, antivirus actions, or any other consequence arising from use of this software or any external tools it downloads or runs.`r`n`r`nLicensed under the Apache License, Version 2.0.`r`nProvided AS IS, without warranties or conditions of any kind.`r`nUse at your own risk."
    $disclaimerGroup.Controls.Add($disclaimerTextBox)

    $aboutTab = New-Object System.Windows.Forms.TabPage
    $aboutTab.Text = 'About'

    $aboutTitleLabel = New-Object System.Windows.Forms.Label
    $aboutTitleLabel.Text = "Ibis v$appVersion"
    $aboutTitleLabel.Font = New-Object System.Drawing.Font($aboutTitleLabel.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $aboutTitleLabel.Location = New-Object System.Drawing.Point(18, 18)
    $aboutTitleLabel.Size = New-Object System.Drawing.Size(620, 34)
    $aboutTab.Controls.Add($aboutTitleLabel)

    $aboutSummaryTextBox = New-Object System.Windows.Forms.TextBox
    $aboutSummaryTextBox.Location = New-Object System.Drawing.Point(18, 66)
    $aboutSummaryTextBox.Size = New-Object System.Drawing.Size(744, 88)
    $aboutSummaryTextBox.Multiline = $true
    $aboutSummaryTextBox.ReadOnly = $true
    $aboutSummaryTextBox.BorderStyle = 'FixedSingle'
    $aboutSummaryTextBox.BackColor = [System.Drawing.SystemColors]::Window
    $aboutSummaryTextBox.Text = "Ibis is currently pre-1.0 beta software. Version numbers are tracked in config.json and release notes are recorded in CHANGELOG.md.`r`n`r`nCurrent version: v$appVersion"
    $aboutTab.Controls.Add($aboutSummaryTextBox)

    $changelogLabel = New-Object System.Windows.Forms.Label
    $changelogLabel.Text = 'Changelog'
    $changelogLabel.Location = New-Object System.Drawing.Point(18, 172)
    $changelogLabel.Size = New-Object System.Drawing.Size(160, 20)
    $aboutTab.Controls.Add($changelogLabel)

    $changelogTextBox = New-Object System.Windows.Forms.TextBox
    $changelogTextBox.Location = New-Object System.Drawing.Point(18, 198)
    $changelogTextBox.Size = New-Object System.Drawing.Size(744, 420)
    $changelogTextBox.Multiline = $true
    $changelogTextBox.ScrollBars = 'Vertical'
    $changelogTextBox.ReadOnly = $true
    $changelogTextBox.BorderStyle = 'FixedSingle'
    $changelogTextBox.BackColor = [System.Drawing.SystemColors]::Window
    if (Test-Path -LiteralPath $changelogPath -PathType Leaf) {
        $changelogTextBox.Text = ConvertTo-IbisGuiDisplayText -Text (Get-Content -LiteralPath $changelogPath -Raw)
    }
    else {
        $changelogTextBox.Text = "CHANGELOG.md was not found at:`r`n$changelogPath"
    }
    $aboutTab.Controls.Add($changelogTextBox)

    $setupTab = New-Object System.Windows.Forms.TabPage
    $setupTab.Text = 'Setup tools'
    $tabs.TabPages.Add($setupTab)

    $toolsLabel = New-Object System.Windows.Forms.Label
    $toolsLabel.Text = 'Tools folder'
    $toolsLabel.Location = New-Object System.Drawing.Point(12, 16)
    $toolsLabel.Width = 180
    $setupTab.Controls.Add($toolsLabel)

    $toolsTextBox = New-Object System.Windows.Forms.TextBox
    $toolsTextBox.Location = New-Object System.Drawing.Point(12, 40)
    $toolsTextBox.Width = 500
    $toolsTextBox.Text = $config.defaultToolsRoot
    $setupTab.Controls.Add($toolsTextBox)

    $toolsBrowseButton = New-Object System.Windows.Forms.Button
    $toolsBrowseButton.Text = 'Browse'
    $toolsBrowseButton.Location = New-Object System.Drawing.Point(526, 38)
    $toolsBrowseButton.Width = 82
    $setupTab.Controls.Add($toolsBrowseButton)

    $openToolsFolderButton = New-Object System.Windows.Forms.Button
    $openToolsFolderButton.Text = 'Open tools folder'
    $openToolsFolderButton.Location = New-Object System.Drawing.Point(620, 38)
    $openToolsFolderButton.Width = 136
    $setupTab.Controls.Add($openToolsFolderButton)

    $toolActionsGroup = New-Object System.Windows.Forms.GroupBox
    $toolActionsGroup.Text = 'Tool management'
    $toolActionsGroup.Location = New-Object System.Drawing.Point(12, 76)
    $toolActionsGroup.Size = New-Object System.Drawing.Size(368, 148)
    $setupTab.Controls.Add($toolActionsGroup)

    $checkToolsButton = New-Object System.Windows.Forms.Button
    $checkToolsButton.Text = 'Recheck Tools'
    $checkToolsButton.Location = New-Object System.Drawing.Point(12, 28)
    $checkToolsButton.Width = 112
    $toolActionsGroup.Controls.Add($checkToolsButton)

    $toolGuidanceButton = New-Object System.Windows.Forms.Button
    $toolGuidanceButton.Text = 'Guidance'
    $toolGuidanceButton.Location = New-Object System.Drawing.Point(12, 68)
    $toolGuidanceButton.Width = 112
    $toolActionsGroup.Controls.Add($toolGuidanceButton)

    $downloadMissingToolsButton = New-Object System.Windows.Forms.Button
    $downloadMissingToolsButton.Text = 'Download Missing Tools'
    $downloadMissingToolsButton.Location = New-Object System.Drawing.Point(136, 28)
    $downloadMissingToolsButton.Width = 214
    $toolActionsGroup.Controls.Add($downloadMissingToolsButton)

    $updateHayabusaRulesButton = New-Object System.Windows.Forms.Button
    $updateHayabusaRulesButton.Text = 'Update Hayabusa Rules'
    $updateHayabusaRulesButton.Location = New-Object System.Drawing.Point(136, 68)
    $updateHayabusaRulesButton.Width = 214
    $toolActionsGroup.Controls.Add($updateHayabusaRulesButton)

    $longPathsLabel = New-Object System.Windows.Forms.Label
    $longPathsLabel.Text = 'Long paths'
    $longPathsLabel.Location = New-Object System.Drawing.Point(12, 110)
    $longPathsLabel.Size = New-Object System.Drawing.Size(82, 20)
    $toolActionsGroup.Controls.Add($longPathsLabel)

    $enableLongPathsButton = New-Object System.Windows.Forms.Button
    $enableLongPathsButton.Text = 'Enable'
    $enableLongPathsButton.Location = New-Object System.Drawing.Point(100, 104)
    $enableLongPathsButton.Width = 82
    $toolActionsGroup.Controls.Add($enableLongPathsButton)

    $disableLongPathsButton = New-Object System.Windows.Forms.Button
    $disableLongPathsButton.Text = 'Disable'
    $disableLongPathsButton.Location = New-Object System.Drawing.Point(192, 104)
    $disableLongPathsButton.Width = 82
    $toolActionsGroup.Controls.Add($disableLongPathsButton)

    $longPathsStatusLabel = New-Object System.Windows.Forms.Label
    $longPathsStatusLabel.Location = New-Object System.Drawing.Point(284, 109)
    $longPathsStatusLabel.Size = New-Object System.Drawing.Size(70, 20)
    $toolActionsGroup.Controls.Add($longPathsStatusLabel)

    $defenderActionsGroup = New-Object System.Windows.Forms.GroupBox
    $defenderActionsGroup.Text = 'Microsoft Defender exclusions'
    $defenderActionsGroup.Location = New-Object System.Drawing.Point(394, 76)
    $defenderActionsGroup.Size = New-Object System.Drawing.Size(362, 96)
    $setupTab.Controls.Add($defenderActionsGroup)

    $defenderAdminLabel = New-Object System.Windows.Forms.Label
    $defenderAdminLabel.Location = New-Object System.Drawing.Point(12, 18)
    $defenderAdminLabel.Size = New-Object System.Drawing.Size(338, 18)
    $defenderAdminLabel.Text = 'Checking administrator permissions...'
    $defenderActionsGroup.Controls.Add($defenderAdminLabel)

    $checkDefenderExclusionsButton = New-Object System.Windows.Forms.Button
    $checkDefenderExclusionsButton.Text = 'Check Existing'
    $checkDefenderExclusionsButton.Location = New-Object System.Drawing.Point(12, 40)
    $checkDefenderExclusionsButton.Width = 108
    $defenderActionsGroup.Controls.Add($checkDefenderExclusionsButton)

    $addDefenderExclusionsButton = New-Object System.Windows.Forms.Button
    $addDefenderExclusionsButton.Text = 'Add'
    $addDefenderExclusionsButton.Location = New-Object System.Drawing.Point(130, 40)
    $addDefenderExclusionsButton.Width = 74
    $defenderActionsGroup.Controls.Add($addDefenderExclusionsButton)

    $removeDefenderExclusionsButton = New-Object System.Windows.Forms.Button
    $removeDefenderExclusionsButton.Text = 'Remove'
    $removeDefenderExclusionsButton.Location = New-Object System.Drawing.Point(214, 40)
    $removeDefenderExclusionsButton.Width = 86
    $defenderActionsGroup.Controls.Add($removeDefenderExclusionsButton)

    $runtimePrereqGroup = New-Object System.Windows.Forms.GroupBox
    $runtimePrereqGroup.Text = 'Runtime prerequisites'
    $runtimePrereqGroup.Location = New-Object System.Drawing.Point(394, 180)
    $runtimePrereqGroup.Size = New-Object System.Drawing.Size(362, 44)
    $setupTab.Controls.Add($runtimePrereqGroup)

    $vcRedistStatusLabel = New-Object System.Windows.Forms.Label
    $vcRedistStatusLabel.Location = New-Object System.Drawing.Point(12, 18)
    $vcRedistStatusLabel.Size = New-Object System.Drawing.Size(214, 18)
    $vcRedistStatusLabel.Text = 'Checking VC++ Redistributable...'
    $runtimePrereqGroup.Controls.Add($vcRedistStatusLabel)

    $openVcRedistPageButton = New-Object System.Windows.Forms.Button
    $openVcRedistPageButton.Text = 'Microsoft page'
    $openVcRedistPageButton.Location = New-Object System.Drawing.Point(236, 14)
    $openVcRedistPageButton.Width = 112
    $runtimePrereqGroup.Controls.Add($openVcRedistPageButton)

    $toolList = New-Object System.Windows.Forms.ListView
    $toolList.Location = New-Object System.Drawing.Point(12, 236)
    $toolList.Size = New-Object System.Drawing.Size(744, 150)
    $toolList.View = 'Details'
    $toolList.FullRowSelect = $true
    $toolList.GridLines = $true
    [void]$toolList.Columns.Add('Status', 90)
    [void]$toolList.Columns.Add('Tool', 220)
    [void]$toolList.Columns.Add('Expected Path', 420)
    $setupTab.Controls.Add($toolList)

    $toolGuidanceTextBox = New-Object System.Windows.Forms.TextBox
    $toolGuidanceTextBox.Location = New-Object System.Drawing.Point(12, 398)
    $toolGuidanceTextBox.Size = New-Object System.Drawing.Size(744, 210)
    $toolGuidanceTextBox.Multiline = $true
    $toolGuidanceTextBox.ScrollBars = 'Vertical'
    $toolGuidanceTextBox.ReadOnly = $true
    $setupTab.Controls.Add($toolGuidanceTextBox)

    $toolHelp = New-Object System.Windows.Forms.Label
    $toolHelp.Location = New-Object System.Drawing.Point(12, 626)
    $toolHelp.Size = New-Object System.Drawing.Size(744, 40)
    $toolHelp.Text = 'Missing tools are not downloaded yet. The guidance view shows where each missing tool is expected and where it can be obtained.'
    $setupTab.Controls.Add($toolHelp)

    $runTab = New-Object System.Windows.Forms.TabPage
    $runTab.Text = 'Run tools'
    $tabs.TabPages.Add($runTab)

    $sourceGroup = New-Object System.Windows.Forms.GroupBox
    $sourceGroup.Text = 'Source'
    $sourceGroup.Location = New-Object System.Drawing.Point(12, 12)
    $sourceGroup.Size = New-Object System.Drawing.Size(744, 112)
    $runTab.Controls.Add($sourceGroup)

    $sourceLabel = New-Object System.Windows.Forms.Label
    $sourceLabel.Text = 'Evidence source root (read-only)'
    $sourceLabel.Location = New-Object System.Drawing.Point(14, 22)
    $sourceLabel.Width = 260
    $sourceGroup.Controls.Add($sourceLabel)

    $sourceTextBox = New-Object System.Windows.Forms.TextBox
    $sourceTextBox.Location = New-Object System.Drawing.Point(14, 46)
    $sourceTextBox.Width = 450
    $sourceTextBox.Text = $config.defaultSourceRoot
    $sourceGroup.Controls.Add($sourceTextBox)

    $sourceBrowseButton = New-Object System.Windows.Forms.Button
    $sourceBrowseButton.Text = 'Browse'
    $sourceBrowseButton.Location = New-Object System.Drawing.Point(476, 44)
    $sourceBrowseButton.Width = 82
    $sourceGroup.Controls.Add($sourceBrowseButton)

    $checkEvidenceButton = New-Object System.Windows.Forms.Button
    $checkEvidenceButton.Text = 'Check source paths exist'
    $checkEvidenceButton.Location = New-Object System.Drawing.Point(568, 44)
    $checkEvidenceButton.Width = 162
    $sourceGroup.Controls.Add($checkEvidenceButton)

    $sourceReadOnlyLabel = New-Object System.Windows.Forms.Label
    $sourceReadOnlyLabel.Text = 'Ibis treats the source as read-only, but cannot control every third-party tool it launches. Best practice is to use a read-only mount, write blocker, or a working copy of the triage pack.'
    $sourceReadOnlyLabel.Location = New-Object System.Drawing.Point(14, 72)
    $sourceReadOnlyLabel.Size = New-Object System.Drawing.Size(716, 34)
    $sourceGroup.Controls.Add($sourceReadOnlyLabel)

    $outputGroup = New-Object System.Windows.Forms.GroupBox
    $outputGroup.Text = 'Output'
    $outputGroup.Location = New-Object System.Drawing.Point(12, 132)
    $outputGroup.Size = New-Object System.Drawing.Size(744, 122)
    $runTab.Controls.Add($outputGroup)

    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Text = 'Output folder'
    $outputLabel.Location = New-Object System.Drawing.Point(14, 22)
    $outputLabel.Width = 260
    $outputGroup.Controls.Add($outputLabel)

    $outputTextBox = New-Object System.Windows.Forms.TextBox
    $outputTextBox.Location = New-Object System.Drawing.Point(14, 46)
    $outputTextBox.Width = 410
    $outputTextBox.Text = $config.defaultOutputRoot
    $outputGroup.Controls.Add($outputTextBox)

    $outputBrowseButton = New-Object System.Windows.Forms.Button
    $outputBrowseButton.Text = 'Browse'
    $outputBrowseButton.Location = New-Object System.Drawing.Point(436, 44)
    $outputBrowseButton.Width = 80
    $outputGroup.Controls.Add($outputBrowseButton)

    $openOutputFolderButton = New-Object System.Windows.Forms.Button
    $openOutputFolderButton.Text = 'Open Output'
    $openOutputFolderButton.Location = New-Object System.Drawing.Point(526, 44)
    $openOutputFolderButton.Width = 110
    $openOutputFolderButton.Enabled = $false
    $outputGroup.Controls.Add($openOutputFolderButton)

    $hostLabel = New-Object System.Windows.Forms.Label
    $hostLabel.Text = 'Hostname prefix'
    $hostLabel.Location = New-Object System.Drawing.Point(14, 76)
    $hostLabel.Width = 104
    $outputGroup.Controls.Add($hostLabel)

    $hostTextBox = New-Object System.Windows.Forms.TextBox
    $hostTextBox.Location = New-Object System.Drawing.Point(118, 74)
    $hostTextBox.Width = 160
    $hostTextBox.Text = $config.defaultHostname
    $outputGroup.Controls.Add($hostTextBox)

    $extractHostNameButton = New-Object System.Windows.Forms.Button
    $extractHostNameButton.Text = 'Extract hostname from SYSTEM hive'
    $extractHostNameButton.Location = New-Object System.Drawing.Point(290, 72)
    $extractHostNameButton.Width = 220
    $outputGroup.Controls.Add($extractHostNameButton)

    $outputWarningLabel = New-Object System.Windows.Forms.Label
    $outputWarningLabel.Location = New-Object System.Drawing.Point(522, 72)
    $outputWarningLabel.Size = New-Object System.Drawing.Size(208, 38)
    $outputWarningLabel.Text = ''
    $outputGroup.Controls.Add($outputWarningLabel)

    $moduleGroup = New-Object System.Windows.Forms.GroupBox
    $moduleGroup.Text = 'Processing modules'
    $moduleGroup.Location = New-Object System.Drawing.Point(12, 262)
    $moduleGroup.Size = New-Object System.Drawing.Size(744, 268)
    $runTab.Controls.Add($moduleGroup)

    $moduleCheckboxes = @()
    $moduleCheckboxById = @{}
    $moduleToolTip = New-Object System.Windows.Forms.ToolTip
    $moduleToolTip.AutoPopDelay = 12000
    $moduleToolTip.InitialDelay = 400
    $moduleToolTip.ReshowDelay = 100
    $moduleToolTip.ShowAlways = $true
    $index = 0
    $rowsPerColumn = [math]::Ceiling($config.modules.Count / 2)
    foreach ($module in $config.modules) {
        $checkBox = New-Object System.Windows.Forms.CheckBox
        $checkBox.Text = $module.name
        $checkBox.Tag = $module
        $checkBox.Checked = [bool]$module.enabledByDefault
        $checkBox.Width = 330
        $column = 0
        if ($index -ge $rowsPerColumn) {
            $column = 1
        }
        $row = $index
        if ($index -ge $rowsPerColumn) {
            $row = $index - $rowsPerColumn
        }
        $checkBox.Location = New-Object System.Drawing.Point((16 + ($column * 360)), (24 + ($row * 22)))
        $moduleGroup.Controls.Add($checkBox)
        $moduleCheckboxes += $checkBox
        $moduleCheckboxById[$module.id] = $checkBox
        if (-not [string]::IsNullOrWhiteSpace([string]$module.hint)) {
            $moduleToolTip.SetToolTip($checkBox, [string]$module.hint)
        }
        $index++
    }

    $selectAllModulesButton = New-Object System.Windows.Forms.Button
    $selectAllModulesButton.Text = 'Select All'
    $selectAllModulesButton.Location = New-Object System.Drawing.Point(16, 226)
    $selectAllModulesButton.Width = 96
    $moduleGroup.Controls.Add($selectAllModulesButton)

    $deselectAllModulesButton = New-Object System.Windows.Forms.Button
    $deselectAllModulesButton.Text = 'Deselect All'
    $deselectAllModulesButton.Location = New-Object System.Drawing.Point(124, 226)
    $deselectAllModulesButton.Width = 104
    $moduleGroup.Controls.Add($deselectAllModulesButton)

    $runProcessingModulesButton = New-Object System.Windows.Forms.Button
    $runProcessingModulesButton.Text = 'Run Selected Modules'
    $runProcessingModulesButton.Location = New-Object System.Drawing.Point(370, 226)
    $runProcessingModulesButton.Width = 150
    $moduleGroup.Controls.Add($runProcessingModulesButton)

    $pauseProcessingButton = New-Object System.Windows.Forms.Button
    $pauseProcessingButton.Text = 'Pause'
    $pauseProcessingButton.Location = New-Object System.Drawing.Point(530, 226)
    $pauseProcessingButton.Width = 82
    $pauseProcessingButton.Enabled = $false
    $moduleGroup.Controls.Add($pauseProcessingButton)

    $cancelProcessingButton = New-Object System.Windows.Forms.Button
    $cancelProcessingButton.Text = 'Cancel'
    $cancelProcessingButton.Location = New-Object System.Drawing.Point(622, 226)
    $cancelProcessingButton.Width = 94
    $cancelProcessingButton.Enabled = $false
    $moduleGroup.Controls.Add($cancelProcessingButton)

    if ($moduleCheckboxById.ContainsKey('hayabusa') -and $moduleCheckboxById.ContainsKey('takajo')) {
        $takajoCheckBox = $moduleCheckboxById['takajo']
        $hayabusaCheckBox = $moduleCheckboxById['hayabusa']
        $takajoCheckBox.Enabled = [bool]$hayabusaCheckBox.Checked
        if (-not $hayabusaCheckBox.Checked) {
            $takajoCheckBox.Checked = $false
        }
        $hayabusaCheckBox.Add_CheckedChanged({
            $takajoCheckBox.Enabled = [bool]$hayabusaCheckBox.Checked
            if (-not $hayabusaCheckBox.Checked) {
                $takajoCheckBox.Checked = $false
            }
        })
    }

    if ($moduleCheckboxById.ContainsKey('eventlogs') -and $moduleCheckboxById.ContainsKey('duckdb-eventlogs')) {
        $duckDbCheckBox = $moduleCheckboxById['duckdb-eventlogs']
        $eventLogsCheckBox = $moduleCheckboxById['eventlogs']
        $duckDbCheckBox.Enabled = [bool]$eventLogsCheckBox.Checked
        if (-not $eventLogsCheckBox.Checked) {
            $duckDbCheckBox.Checked = $false
        }
        $eventLogsCheckBox.Add_CheckedChanged({
            $duckDbCheckBox.Enabled = [bool]$eventLogsCheckBox.Checked
            if (-not $eventLogsCheckBox.Checked) {
                $duckDbCheckBox.Checked = $false
            }
        })
    }

    $runProgressLabel = New-Object System.Windows.Forms.Label
    $runProgressLabel.Text = 'Ready'
    $runProgressLabel.Location = New-Object System.Drawing.Point(12, 534)
    $runProgressLabel.Size = New-Object System.Drawing.Size(350, 34)
    $runTab.Controls.Add($runProgressLabel)

    $runProgressBar = New-Object System.Windows.Forms.ProgressBar
    $runProgressBar.Location = New-Object System.Drawing.Point(376, 542)
    $runProgressBar.Size = New-Object System.Drawing.Size(380, 18)
    $runProgressBar.Style = 'Continuous'
    $runProgressBar.Minimum = 0
    $runProgressBar.Maximum = 100
    $runProgressBar.Value = 0
    $runProgressBar.MarqueeAnimationSpeed = 0
    $runProgressBar.Visible = $false
    $runTab.Controls.Add($runProgressBar)

    $logTextBox = New-Object System.Windows.Forms.TextBox
    $logTextBox.Location = New-Object System.Drawing.Point(12, 574)
    $logTextBox.Size = New-Object System.Drawing.Size(744, 130)
    $logTextBox.Multiline = $true
    $logTextBox.ScrollBars = 'Vertical'
    $logTextBox.ReadOnly = $true
    $runTab.Controls.Add($logTextBox)

    $settingsTab = New-Object System.Windows.Forms.TabPage
    $settingsTab.Text = 'Settings'
    $tabs.TabPages.Add($settingsTab)

    $settingsTitleLabel = New-Object System.Windows.Forms.Label
    $settingsTitleLabel.Text = 'Settings'
    $settingsTitleLabel.Font = New-Object System.Drawing.Font($settingsTitleLabel.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $settingsTitleLabel.Location = New-Object System.Drawing.Point(18, 18)
    $settingsTitleLabel.Size = New-Object System.Drawing.Size(720, 26)
    $settingsTab.Controls.Add($settingsTitleLabel)

    $notificationGroup = New-Object System.Windows.Forms.GroupBox
    $notificationGroup.Text = 'Notifications'
    $notificationGroup.Location = New-Object System.Drawing.Point(18, 58)
    $notificationGroup.Size = New-Object System.Drawing.Size(744, 112)
    $settingsTab.Controls.Add($notificationGroup)

    $completionBeepCheckBox = New-Object System.Windows.Forms.CheckBox
    $completionBeepCheckBox.Text = 'Play audible beep when processing run completes'
    $completionBeepCheckBox.Location = New-Object System.Drawing.Point(16, 30)
    $completionBeepCheckBox.Size = New-Object System.Drawing.Size(430, 24)
    $completionBeepCheckBox.Checked = $completionBeepEnabled
    $notificationGroup.Controls.Add($completionBeepCheckBox)

    $notificationInfoLabel = New-Object System.Windows.Forms.Label
    $notificationInfoLabel.Text = 'Ibis will still show a completion popup even when the audible beep is disabled.'
    $notificationInfoLabel.Location = New-Object System.Drawing.Point(16, 62)
    $notificationInfoLabel.Size = New-Object System.Drawing.Size(700, 32)
    $notificationGroup.Controls.Add($notificationInfoLabel)

    $logsTab = New-Object System.Windows.Forms.TabPage
    $logsTab.Text = 'Logs'
    $tabs.TabPages.Add($logsTab)

    $logsTitleLabel = New-Object System.Windows.Forms.Label
    $logsTitleLabel.Text = 'Session logging'
    $logsTitleLabel.Font = New-Object System.Drawing.Font($logsTitleLabel.Font.FontFamily, 12, [System.Drawing.FontStyle]::Bold)
    $logsTitleLabel.Location = New-Object System.Drawing.Point(18, 18)
    $logsTitleLabel.Size = New-Object System.Drawing.Size(720, 26)
    $logsTab.Controls.Add($logsTitleLabel)

    $logsInfoTextBox = New-Object System.Windows.Forms.TextBox
    $logsInfoTextBox.Location = New-Object System.Drawing.Point(18, 54)
    $logsInfoTextBox.Size = New-Object System.Drawing.Size(744, 92)
    $logsInfoTextBox.Multiline = $true
    $logsInfoTextBox.ReadOnly = $true
    $logsInfoTextBox.BorderStyle = 'FixedSingle'
    $logsInfoTextBox.BackColor = [System.Drawing.SystemColors]::Window
    $logsInfoTextBox.Text = "Ibis creates one log file for each GUI session. The log records setup actions, processing messages, failures, and command line hints for forensic tools so commands can be copied and rerun manually for debugging."
    $logsTab.Controls.Add($logsInfoTextBox)

    $logsDirectoryLabel = New-Object System.Windows.Forms.Label
    $logsDirectoryLabel.Text = 'Logs directory'
    $logsDirectoryLabel.Location = New-Object System.Drawing.Point(18, 166)
    $logsDirectoryLabel.Width = 160
    $logsTab.Controls.Add($logsDirectoryLabel)

    $logsDirectoryTextBox = New-Object System.Windows.Forms.TextBox
    $logsDirectoryTextBox.Location = New-Object System.Drawing.Point(18, 190)
    $logsDirectoryTextBox.Size = New-Object System.Drawing.Size(610, 22)
    $logsDirectoryTextBox.ReadOnly = $true
    $logsDirectoryTextBox.Text = $logsDirectory
    $logsTab.Controls.Add($logsDirectoryTextBox)

    $openLogsDirectoryButton = New-Object System.Windows.Forms.Button
    $openLogsDirectoryButton.Text = 'Open Logs Folder'
    $openLogsDirectoryButton.Location = New-Object System.Drawing.Point(644, 188)
    $openLogsDirectoryButton.Width = 118
    $logsTab.Controls.Add($openLogsDirectoryButton)

    $sessionLogLabel = New-Object System.Windows.Forms.Label
    $sessionLogLabel.Text = 'Current session log'
    $sessionLogLabel.Location = New-Object System.Drawing.Point(18, 232)
    $sessionLogLabel.Width = 180
    $logsTab.Controls.Add($sessionLogLabel)

    $sessionLogTextBox = New-Object System.Windows.Forms.TextBox
    $sessionLogTextBox.Location = New-Object System.Drawing.Point(18, 256)
    $sessionLogTextBox.Size = New-Object System.Drawing.Size(610, 22)
    $sessionLogTextBox.ReadOnly = $true
    $sessionLogTextBox.Text = $sessionLogPath
    $logsTab.Controls.Add($sessionLogTextBox)

    $openSessionLogButton = New-Object System.Windows.Forms.Button
    $openSessionLogButton.Text = 'Open Current Log'
    $openSessionLogButton.Location = New-Object System.Drawing.Point(644, 254)
    $openSessionLogButton.Width = 118
    $logsTab.Controls.Add($openSessionLogButton)

    $tabs.TabPages.Add($aboutTab)

    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $downloadState = @{
        Operation = $null
        ProgressPath = $null
        ProgressLineCount = 0
        FileSnapshot = $null
    }
    $downloadPollTimer = New-Object System.Windows.Forms.Timer
    $downloadPollTimer.Interval = 750
    $hayabusaRulesUpdateState = @{
        Operation = $null
    }
    $hayabusaRulesUpdatePollTimer = New-Object System.Windows.Forms.Timer
    $hayabusaRulesUpdatePollTimer.Interval = 750
    $processingState = @{
        Operation = $null
        ProgressPath = $null
        ProgressLineCount = 0
        ControlPath = $null
        IsPaused = $false
        CancelRequested = $false
    }
    $processingPollTimer = New-Object System.Windows.Forms.Timer
    $processingPollTimer.Interval = 750

    $appendDownloadProgress = {
        if ([string]::IsNullOrWhiteSpace($downloadState.ProgressPath)) {
            return
        }

        $progress = Read-IbisProgressEvents -ProgressPath $downloadState.ProgressPath -SkipCount $downloadState.ProgressLineCount
        $downloadState.ProgressLineCount = $progress.LineCount
        foreach ($progressEvent in $progress.Events) {
            $prefix = ''
            if ($progressEvent.Total -gt 0 -and $progressEvent.Index -gt 0) {
                $prefix = "[$($progressEvent.Index)/$($progressEvent.Total)] "
            }

            $name = $progressEvent.ToolName
            if ([string]::IsNullOrWhiteSpace($name)) {
                $name = 'Ibis'
            }

            $message = "${prefix}${name}: $($progressEvent.Stage) - $($progressEvent.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$message`r`n" -StripAnsi
            $statusLabel.Text = $message
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $message
        }
    }

    $appendProcessingProgress = {
        if ([string]::IsNullOrWhiteSpace($processingState.ProgressPath)) {
            return
        }

        $progress = Read-IbisProgressEvents -ProgressPath $processingState.ProgressPath -SkipCount $processingState.ProgressLineCount
        $processingState.ProgressLineCount = $progress.LineCount
        foreach ($progressEvent in $progress.Events) {
            $prefix = ''
            if ($progressEvent.Total -gt 0 -and $progressEvent.Index -gt 0) {
                $prefix = "[$($progressEvent.Index)/$($progressEvent.Total)] "
            }

            $name = $progressEvent.ToolName
            if ([string]::IsNullOrWhiteSpace($name)) {
                $name = 'Ibis'
            }

            $message = "${prefix}${name}: $($progressEvent.Stage) - $($progressEvent.Message)"
            if ($progressEvent.Status -eq 'Audit') {
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $message
                continue
            }

            Add-IbisLogLine -LogTextBox $logTextBox -Message $message
            $statusLabel.Text = $message
            $runProgressLabel.Text = $message
            if ($progressEvent.Total -gt 0) {
                $runProgressBar.Style = 'Continuous'
                $runProgressBar.Maximum = [math]::Max(1, [int]$progressEvent.Total)
                $runProgressBar.Value = [math]::Min([int]$progressEvent.Index, $runProgressBar.Maximum)
                $runProgressBar.Visible = $true
            }
        }
    }

    $updateToolsFolderButtonState = {
        $openToolsFolderButton.Enabled = (-not [string]::IsNullOrWhiteSpace($toolsTextBox.Text) -and (Test-Path -LiteralPath $toolsTextBox.Text -PathType Container))
    }

    $updateLongPathsControls = {
        $isAdministrator = Test-IbisIsAdministrator
        $longPathsEnabled = Get-IbisLongPathsEnabled
        $enableLongPathsButton.Enabled = ($isAdministrator -and -not $longPathsEnabled)
        $disableLongPathsButton.Enabled = ($isAdministrator -and $longPathsEnabled)
        if ($longPathsEnabled) {
            $longPathsStatusLabel.Text = 'Enabled'
            $longPathsStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
        else {
            $longPathsStatusLabel.Text = 'Disabled'
            $longPathsStatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    }

    $updateVcRedistStatus = {
        $vcStatus = Get-IbisVisualCppRedistributableStatus -Architecture 'x64'
        if ($vcStatus.Present) {
            $vcRedistStatusLabel.Text = 'VC++ 2015+ x64 detected'
            $vcRedistStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
        else {
            $vcRedistStatusLabel.Text = 'VC++ 2015+ x64 missing'
            $vcRedistStatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        $vcStatus
    }

    $setDefenderControlsForAdmin = {
        $isAdministrator = Test-IbisIsAdministrator
        $checkDefenderExclusionsButton.Enabled = $isAdministrator
        $addDefenderExclusionsButton.Enabled = $isAdministrator
        $removeDefenderExclusionsButton.Enabled = $isAdministrator
        & $updateLongPathsControls
        if ($isAdministrator) {
            $defenderAdminLabel.Text = 'Administrator permissions detected.'
            $defenderAdminLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
        else {
            $defenderAdminLabel.Text = 'No administrator permissions detected.'
            $defenderAdminLabel.ForeColor = [System.Drawing.Color]::DarkRed
        }
    }

    & $setDefenderControlsForAdmin
    & $updateToolsFolderButtonState
    [void](& $updateVcRedistStatus)

    $updateModuleDependencies = {
        if ($moduleCheckboxById.ContainsKey('hayabusa') -and $moduleCheckboxById.ContainsKey('takajo')) {
            $moduleCheckboxById['takajo'].Enabled = [bool]$moduleCheckboxById['hayabusa'].Checked
            if (-not $moduleCheckboxById['hayabusa'].Checked) {
                $moduleCheckboxById['takajo'].Checked = $false
            }
        }
        if ($moduleCheckboxById.ContainsKey('eventlogs') -and $moduleCheckboxById.ContainsKey('duckdb-eventlogs')) {
            $moduleCheckboxById['duckdb-eventlogs'].Enabled = [bool]$moduleCheckboxById['eventlogs'].Checked
            if (-not $moduleCheckboxById['eventlogs'].Checked) {
                $moduleCheckboxById['duckdb-eventlogs'].Checked = $false
            }
        }
    }

    $updateOutputWarning = {
        $path = $outputTextBox.Text
        if ([string]::IsNullOrWhiteSpace($path)) {
            $outputWarningLabel.Text = 'Select an output folder.'
            $outputWarningLabel.ForeColor = [System.Drawing.Color]::DarkRed
            $openOutputFolderButton.Enabled = $false
            return
        }

        if (-not [string]::IsNullOrWhiteSpace($sourceTextBox.Text) -and (Test-IbisPathInsideRoot -Path $path -RootPath $sourceTextBox.Text)) {
            $outputWarningLabel.Text = 'Output is inside the source. Choose a separate output folder.'
            $outputWarningLabel.ForeColor = [System.Drawing.Color]::DarkRed
            $openOutputFolderButton.Enabled = (Test-Path -LiteralPath $path -PathType Container)
            return
        }

        if (-not (Test-Path -LiteralPath $path -PathType Container)) {
            $outputWarningLabel.Text = 'Output folder does not exist; it will be created.'
            $outputWarningLabel.ForeColor = [System.Drawing.Color]::DarkOrange
            $openOutputFolderButton.Enabled = $false
            return
        }

        $openOutputFolderButton.Enabled = $true
        $existingItems = @(Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($existingItems.Count -gt 0) {
            $outputWarningLabel.Text = 'Output folder already contains data; files may be overwritten.'
            $outputWarningLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        else {
            $outputWarningLabel.Text = 'Output folder exists and is empty.'
            $outputWarningLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        }
    }

    $refreshToolStatusList = {
        & $updateToolsFolderButtonState
        $vcStatus = & $updateVcRedistStatus
        $toolList.Items.Clear()
        $statuses = @(Test-IbisToolStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        foreach ($status in $statuses) {
            $item = New-Object System.Windows.Forms.ListViewItem($status.Status)
            [void]$item.SubItems.Add($status.Name)
            [void]$item.SubItems.Add($status.ExpectedPath)
            if ($status.Present) {
                $item.ForeColor = [System.Drawing.Color]::DarkGreen
            }
            else {
                $item.ForeColor = [System.Drawing.Color]::DarkRed
            }
            [void]$toolList.Items.Add($item)
        }

        $missing = @($statuses | Where-Object { -not $_.Present })
        $statusLabel.Text = "$($statuses.Count) tools checked; $($missing.Count) missing"
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $vcStatus.Message
        if (-not $vcStatus.Present) {
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Warning: $($vcStatus.Message)`r`nInstall with winget if available: winget install -e --id Microsoft.VCRedist.2015+.x64`r`n" -StripAnsi
        }
    }

    $testSourceWriteBoundary = {
        $writablePaths = @(
            [pscustomobject]@{ Name = 'Output folder'; Path = $outputTextBox.Text },
            [pscustomobject]@{ Name = 'Tools folder'; Path = $toolsTextBox.Text },
            [pscustomobject]@{ Name = 'Logs folder'; Path = $logsDirectory },
            [pscustomobject]@{ Name = 'Temporary working folder'; Path = [System.IO.Path]::GetTempPath() }
        )

        Test-IbisSourceWriteBoundary -SourceRoot $sourceTextBox.Text -WritablePaths $writablePaths
    }

    $stopIfSourceWriteBoundaryFails = {
        param(
            [string]$ActionName = 'operation'
        )

        $boundaryResult = & $testSourceWriteBoundary
        if ($boundaryResult.Passed) {
            Add-IbisLogLine -LogTextBox $logTextBox -Message "Source read-only boundary check passed for ${ActionName}: writable folders are outside the evidence source."
            return $false
        }

        Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Source read-only boundary check failed for ${ActionName}. Ibis will not start because a writable folder is inside the evidence source."
        foreach ($violation in $boundaryResult.Violations) {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "$($violation.Name) is inside the source: $($violation.Path)"
        }
        $statusLabel.Text = 'Source read-only boundary check failed'
        [System.Windows.Forms.MessageBox]::Show(
            "Ibis will not start this ${ActionName} because one or more writable folders are inside the selected evidence source.`r`n`r`nChoose output, tools, logs, and temporary folders outside the evidence source.",
            'Source Read-Only Boundary',
            'OK',
            'Error'
        ) | Out-Null
        $true
    }

    $saveConfigPaths = {
        param(
            [string]$Reason = 'path update'
        )

        try {
            [void](Save-IbisConfigPathSetting `
                -ProjectRoot $ProjectRoot `
                -ToolsRoot $toolsTextBox.Text `
                -SourceRoot $sourceTextBox.Text `
                -OutputRoot $outputTextBox.Text `
                -CompletionBeepEnabled ([bool]$completionBeepCheckBox.Checked))
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Saved configuration after ${Reason}: tools='$($toolsTextBox.Text)', source='$($sourceTextBox.Text)', output='$($outputTextBox.Text)', completionBeepEnabled=$([bool]$completionBeepCheckBox.Checked)"
        }
        catch {
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'WARN' -Message "Unable to save configuration after ${Reason}: $($_.Exception.Message)"
        }
    }

    $showProcessingCompletionNotification = {
        param(
            [string]$Message,
            [string]$Title = 'Ibis Processing Complete',
            [string]$Icon = 'Information'
        )

        if ([bool]$completionBeepCheckBox.Checked) {
            try {
                [System.Media.SystemSounds]::Asterisk.Play()
            }
            catch {
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'WARN' -Message "Unable to play completion beep: $($_.Exception.Message)"
            }
        }

        $messageBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Information
        if ($Icon -eq 'Warning') {
            $messageBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Warning
        }
        elseif ($Icon -eq 'Error') {
            $messageBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Error
        }

        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $messageBoxIcon
        ) | Out-Null
    }

    & $updateOutputWarning

    $completionBeepCheckBox.Add_CheckedChanged({
        & $saveConfigPaths -Reason 'completion notification setting update'
        $statusLabel.Text = 'Settings saved'
    })

    $openLogsDirectoryButton.Add_Click({
        if (-not (Test-Path -LiteralPath $logsDirectory)) {
            New-Item -ItemType Directory -Path $logsDirectory -Force | Out-Null
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Opening logs directory: $logsDirectory"
        Start-Process -FilePath 'explorer.exe' -ArgumentList $logsDirectory
    })

    $openSessionLogButton.Add_Click({
        if (-not (Test-Path -LiteralPath $sessionLogPath -PathType Leaf)) {
            New-Item -ItemType File -Path $sessionLogPath -Force | Out-Null
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Opening current session log: $sessionLogPath"
        Start-Process -FilePath $sessionLogPath
    })

    $toolsBrowseButton.Add_Click({
        $folderDialog.SelectedPath = $toolsTextBox.Text
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $toolsTextBox.Text = $folderDialog.SelectedPath
            & $updateToolsFolderButtonState
            & $saveConfigPaths -Reason 'tools folder selection'
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Tools folder selected: $($toolsTextBox.Text)"
        }
    })

    $openToolsFolderButton.Add_Click({
        if (Test-Path -LiteralPath $toolsTextBox.Text -PathType Container) {
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Opening tools folder: $($toolsTextBox.Text)"
            Start-Process -FilePath 'explorer.exe' -ArgumentList @($toolsTextBox.Text)
        }
    })

    $toolsTextBox.Add_TextChanged({
        & $updateToolsFolderButtonState
    })

    $enableLongPathsButton.Add_Click({
        try {
            $result = Set-IbisLongPathsEnabled -Enabled $true
            & $updateLongPathsControls
            $message = "Windows long path support: $($result.Message)"
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text $message -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $message
            $statusLabel.Text = 'Windows long path support enabled'
        }
        catch {
            $message = "Unable to enable Windows long path support: $($_.Exception.Message)"
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text $message -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $message
            $statusLabel.Text = 'Long path update failed'
        }
    })

    $disableLongPathsButton.Add_Click({
        try {
            $result = Set-IbisLongPathsEnabled -Enabled $false
            & $updateLongPathsControls
            $message = "Windows long path support: $($result.Message)"
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text $message -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $message
            $statusLabel.Text = 'Windows long path support disabled'
        }
        catch {
            $message = "Unable to disable Windows long path support: $($_.Exception.Message)"
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text $message -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $message
            $statusLabel.Text = 'Long path update failed'
        }
    })

    $openVcRedistPageButton.Add_Click({
        $vcStatus = Get-IbisVisualCppRedistributableStatus -Architecture 'x64'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Opening Visual C++ Redistributable page: $($vcStatus.MicrosoftUrl)"
        Start-Process -FilePath $vcStatus.MicrosoftUrl
    })

    $sourceBrowseButton.Add_Click({
        $folderDialog.SelectedPath = $sourceTextBox.Text
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $sourceTextBox.Text = $folderDialog.SelectedPath
            & $updateOutputWarning
            & $saveConfigPaths -Reason 'source folder selection'
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Source folder selected: $($sourceTextBox.Text)"
        }
    })

    $outputBrowseButton.Add_Click({
        $folderDialog.SelectedPath = $outputTextBox.Text
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputTextBox.Text = $folderDialog.SelectedPath
            & $updateOutputWarning
            & $saveConfigPaths -Reason 'output folder selection'
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Output folder selected: $($outputTextBox.Text)"
        }
    })

    $openOutputFolderButton.Add_Click({
        $path = $outputTextBox.Text
        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Container)) {
            $openOutputFolderButton.Enabled = $false
            $statusLabel.Text = 'Output folder does not exist yet'
            return
        }

        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Opening output folder: $path"
        Start-Process -FilePath 'explorer.exe' -ArgumentList $path
    })

    $outputTextBox.Add_TextChanged({
        & $updateOutputWarning
    })

    $sourceTextBox.Add_TextChanged({
        & $updateOutputWarning
    })

    $selectAllModulesButton.Add_Click({
        foreach ($checkBox in $moduleCheckboxes) {
            $checkBox.Checked = $true
        }
        & $updateModuleDependencies
        $statusLabel.Text = 'All processing modules selected'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $deselectAllModulesButton.Add_Click({
        foreach ($checkBox in $moduleCheckboxes) {
            $checkBox.Checked = $false
        }
        & $updateModuleDependencies
        $statusLabel.Text = 'All processing modules deselected'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $checkToolsButton.Add_Click({
        & $refreshToolStatusList
    })

    $toolGuidanceButton.Add_Click({
        $statuses = @(Test-IbisToolStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        $plan = @(Get-IbisToolAcquisitionPlan -ToolStatuses $statuses)
        if ($plan.Count -eq 0) {
            $plan = @($statuses | ForEach-Object {
                $source = $_.DownloadUrl
                if ([string]::IsNullOrWhiteSpace($source) -or $source -eq 'latest-release') {
                    $source = $_.ManualUrl
                }

                [pscustomobject]@{
                    Id = $_.Id
                    Name = $_.Name
                    ExpectedPath = $_.ExpectedPath
                    AcquisitionSource = $source
                    Notes = $_.Notes
                }
            })
        }
        Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text (Format-IbisToolAcquisitionPlan -AcquisitionPlan $plan) -StripAnsi
        $missing = @($statuses | Where-Object { -not $_.Present })
        if ($missing.Count -eq 0) {
            $statusLabel.Text = "$($plan.Count) tool guidance entries"
        }
        else {
            $statusLabel.Text = "$($plan.Count) missing tool guidance entries"
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $checkDefenderExclusionsButton.Add_Click({
        $statuses = @(Get-IbisDefenderExclusionStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        $toolGuidanceTextBox.Clear()
        if ($statuses.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'No Defender exclusions are recommended by the current tool configuration.' -StripAnsi
            $statusLabel.Text = 'No Defender exclusions configured'
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
            return
        }

        if (-not (Test-IbisIsAdministrator)) {
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Warning: Ibis is not running as Administrator. Defender exclusion checks may be incomplete or denied by Windows. Run as Administrator for an authoritative result.`r`n`r`n" -StripAnsi
        }

        foreach ($status in $statuses) {
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$($status.Name): $($status.Status) - $($status.Path)`r`n" -StripAnsi
            if ($status.IsAuthoritative -eq $false) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "  Check is not authoritative in this session.`r`n" -StripAnsi
            }
            if (-not [string]::IsNullOrWhiteSpace($status.Reason)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "  Reason: $($status.Reason)`r`n" -StripAnsi
            }
            if (-not [string]::IsNullOrWhiteSpace($status.Message)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "  $($status.Message)`r`n" -StripAnsi
            }
        }

        $missing = @($statuses | Where-Object { -not $_.Present })
        if (-not (Test-IbisIsAdministrator)) {
            $statusLabel.Text = "$($statuses.Count) exclusions checked; standard-user result may be incomplete"
        }
        else {
            $statusLabel.Text = "$($statuses.Count) exclusions checked; $($missing.Count) missing"
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $addDefenderExclusionsButton.Add_Click({
        $statuses = @(Get-IbisDefenderExclusionStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        $missing = @($statuses | Where-Object { -not $_.Present -and $_.Status -ne 'Unavailable' -and $_.Status -ne 'Failed' })
        if ($statuses.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'No Defender exclusions are recommended by the current tool configuration.' -StripAnsi
            $statusLabel.Text = 'No Defender exclusions configured'
            return
        }

        if (-not (Test-IbisIsAdministrator)) {
            $toolGuidanceTextBox.Clear()
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'Ibis is not running as Administrator. Windows may not allow Defender exclusions to be checked or added from this session. Re-run PowerShell as Administrator and try again.' -StripAnsi
            $statusLabel.Text = 'Administrator required for Defender exclusions'
            return
        }

        if ($missing.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'All recommended Defender exclusions are already present, or Defender is unavailable.' -StripAnsi
            $statusLabel.Text = 'No missing Defender exclusions'
            return
        }

        $paths = ($missing | ForEach-Object { $_.Path }) -join "`r`n"
        $message = "Ibis will add Windows Defender folder exclusions for:`r`n`r`n$paths`r`n`r`nThis may require running PowerShell as Administrator. Continue?"
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Add Defender Exclusions', 'YesNo', 'Warning')
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = 'Defender exclusion update cancelled'
            return
        }

        $toolGuidanceTextBox.Clear()
        $results = @(Add-IbisDefenderExclusion -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        foreach ($result in $results) {
            $resultMessage = "$($result.Name): $($result.Status) - $($result.Path)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$resultMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $resultMessage
            if (-not [string]::IsNullOrWhiteSpace($result.Message)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "  $($result.Message)`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "  $($result.Message)"
            }
        }

        $failed = @($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Unavailable' })
        if ($failed.Count -gt 0) {
            $statusLabel.Text = "Defender exclusions attempted; $($failed.Count) need review"
        }
        else {
            $statusLabel.Text = 'Defender exclusions updated'
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $removeDefenderExclusionsButton.Add_Click({
        $statuses = @(Get-IbisDefenderExclusionStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        $present = @($statuses | Where-Object { $_.Present -and $_.Status -ne 'Unavailable' -and $_.Status -ne 'Failed' })
        if ($statuses.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'No Defender exclusions are recommended by the current tool configuration.' -StripAnsi
            $statusLabel.Text = 'No Defender exclusions configured'
            return
        }

        if (-not (Test-IbisIsAdministrator)) {
            & $setDefenderControlsForAdmin
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'Ibis is not running as Administrator. Windows may not allow Defender exclusions to be removed from this session. Re-run PowerShell as Administrator and try again.' -StripAnsi
            $statusLabel.Text = 'Administrator required for Defender exclusions'
            return
        }

        if ($present.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'No configured Defender exclusions are currently present, or Defender is unavailable.' -StripAnsi
            $statusLabel.Text = 'No Defender exclusions to remove'
            return
        }

        $paths = ($present | ForEach-Object { $_.Path }) -join "`r`n"
        $message = "Ibis will remove Windows Defender folder exclusions for:`r`n`r`n$paths`r`n`r`nContinue?"
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Remove Defender Exclusions', 'YesNo', 'Warning')
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = 'Defender exclusion removal cancelled'
            return
        }

        $toolGuidanceTextBox.Clear()
        $results = @(Remove-IbisDefenderExclusion -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        foreach ($result in $results) {
            $resultMessage = "$($result.Name): $($result.Status) - $($result.Path)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$resultMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $resultMessage
            if (-not [string]::IsNullOrWhiteSpace($result.Message)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "  $($result.Message)`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "  $($result.Message)"
            }
        }

        $failed = @($results | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 'Unavailable' })
        if ($failed.Count -gt 0) {
            $statusLabel.Text = "Defender exclusion removal attempted; $($failed.Count) need review"
        }
        else {
            $statusLabel.Text = 'Defender exclusions removed'
        }
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
    })

    $downloadPollTimer.Add_Tick({
        if ($null -eq $downloadState.Operation) {
            $downloadPollTimer.Stop()
            return
        }

        & $appendDownloadProgress

        if (-not $downloadState.Operation.Handle.IsCompleted) {
            $elapsed = [int]((Get-Date) - $downloadState.Operation.Started).TotalSeconds
            if ([string]::IsNullOrWhiteSpace($statusLabel.Text) -or $statusLabel.Text -eq 'Downloading missing tools') {
                $statusLabel.Text = "Downloading missing tools... ${elapsed}s"
            }
            return
        }

        $downloadPollTimer.Stop()
        try {
            & $appendDownloadProgress
            $results = @($downloadState.Operation.PowerShell.EndInvoke($downloadState.Operation.Handle))
            foreach ($errorRecord in $downloadState.Operation.PowerShell.Streams.Error) {
                $errorMessage = "Error: $($errorRecord.Exception.Message)"
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$errorMessage`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $errorMessage
            }

            foreach ($result in $results) {
                $resultMessage = "$($result.Name): $($result.Status) - $($result.Message)"
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$resultMessage`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $resultMessage
            }

            $toolList.Items.Clear()
            $updatedStatuses = @(Test-IbisToolStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
            foreach ($status in $updatedStatuses) {
                $item = New-Object System.Windows.Forms.ListViewItem($status.Status)
                [void]$item.SubItems.Add($status.Name)
                [void]$item.SubItems.Add($status.ExpectedPath)
                if ($status.Present) {
                    $item.ForeColor = [System.Drawing.Color]::DarkGreen
                }
                else {
                    $item.ForeColor = [System.Drawing.Color]::DarkRed
                }
                [void]$toolList.Items.Add($item)
            }

            $remainingMissing = @($updatedStatuses | Where-Object { -not $_.Present })
            $statusLabel.Text = "Download complete; $($remainingMissing.Count) tools still missing"
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $statusLabel.Text
        }
        catch {
            $downloadErrorMessage = "Download/install failed: $($_.Exception.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$downloadErrorMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $downloadErrorMessage
            $statusLabel.Text = 'Tool download failed'
        }
        finally {
            Stop-IbisToolInstallRunspace -Operation $downloadState.Operation
            if (-not [string]::IsNullOrWhiteSpace($downloadState.ProgressPath) -and (Test-Path -LiteralPath $downloadState.ProgressPath)) {
                try {
                    Remove-Item -LiteralPath $downloadState.ProgressPath -Force
                }
                catch {
                }
            }
            if ($null -ne $downloadState.FileSnapshot) {
                $afterDownloadSnapshot = Get-IbisFileSystemSnapshot -RootPath $downloadState.FileSnapshot.RootPath
                Add-IbisFileSystemChangeLog -Before $downloadState.FileSnapshot -After $afterDownloadSnapshot -Context 'tool download/install' -LogFilePath $sessionLogPath
            }
            $downloadState.Operation = $null
            $downloadState.ProgressPath = $null
            $downloadState.ProgressLineCount = 0
            $downloadState.FileSnapshot = $null
            $downloadMissingToolsButton.Enabled = $true
            $updateHayabusaRulesButton.Enabled = $true
            $checkToolsButton.Enabled = $true
            $toolGuidanceButton.Enabled = $true
            & $setDefenderControlsForAdmin
            $toolsBrowseButton.Enabled = $true
            $toolsTextBox.Enabled = $true
            & $updateToolsFolderButtonState
        }
    })

    $downloadMissingToolsButton.Add_Click({
        if ($null -ne $downloadState.Operation) {
            $statusLabel.Text = 'Tool download already running'
            return
        }

        $statuses = @(Test-IbisToolStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions)
        $missing = @($statuses | Where-Object { -not $_.Present })
        if ($missing.Count -eq 0) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'All configured tools are present.' -StripAnsi
            $statusLabel.Text = 'No missing tools'
            return
        }

        $message = "Ibis will attempt to download and install $($missing.Count) missing tools into:`r`n`r`n$($toolsTextBox.Text)`r`n`r`nSome tools may be flagged by antivirus products. Continue?"
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Download Missing Tools', 'YesNo', 'Warning')
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = 'Tool download cancelled'
            return
        }

        $downloadMissingToolsButton.Enabled = $false
        $updateHayabusaRulesButton.Enabled = $false
        $checkToolsButton.Enabled = $false
        $toolGuidanceButton.Enabled = $false
        $checkDefenderExclusionsButton.Enabled = $false
        $addDefenderExclusionsButton.Enabled = $false
        $removeDefenderExclusionsButton.Enabled = $false
        $toolsBrowseButton.Enabled = $false
        $openToolsFolderButton.Enabled = $false
        $toolsTextBox.Enabled = $false
        $toolGuidanceTextBox.Clear()
        Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Starting background download/install for $($missing.Count) missing tools...`r`n" -StripAnsi
        Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "The GUI should remain responsive while this runs. Progress will appear here as each tool moves through download, extract, publish, and post-install checks.`r`n" -StripAnsi
        $statusLabel.Text = 'Downloading missing tools'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Starting background download/install for $($missing.Count) missing tools into $($toolsTextBox.Text)."
        $downloadState.FileSnapshot = Get-IbisFileSystemSnapshot -RootPath $toolsTextBox.Text

        try {
            $downloadState.ProgressPath = Join-Path ([System.IO.Path]::GetTempPath()) ('IbisToolInstallProgress-' + [System.Guid]::NewGuid().ToString() + '.jsonl')
            $downloadState.ProgressLineCount = 0
            $downloadState.Operation = Start-IbisToolInstallRunspace -ProjectRoot $ProjectRoot -ToolsRoot $toolsTextBox.Text -ProgressPath $downloadState.ProgressPath
            $downloadPollTimer.Start()
        }
        catch {
            $downloadStartErrorMessage = "Unable to start background download/install: $($_.Exception.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$downloadStartErrorMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $downloadStartErrorMessage
            $statusLabel.Text = 'Tool download failed to start'
            if ($null -ne $downloadState.Operation) {
                Stop-IbisToolInstallRunspace -Operation $downloadState.Operation
                $downloadState.Operation = $null
            }
            if (-not [string]::IsNullOrWhiteSpace($downloadState.ProgressPath) -and (Test-Path -LiteralPath $downloadState.ProgressPath)) {
                try {
                    Remove-Item -LiteralPath $downloadState.ProgressPath -Force
                }
                catch {
                }
            }
            $downloadState.ProgressPath = $null
            $downloadState.ProgressLineCount = 0
            $downloadState.FileSnapshot = $null
            $downloadMissingToolsButton.Enabled = $true
            $updateHayabusaRulesButton.Enabled = $true
            $checkToolsButton.Enabled = $true
            $toolGuidanceButton.Enabled = $true
            & $setDefenderControlsForAdmin
            $toolsBrowseButton.Enabled = $true
            $toolsTextBox.Enabled = $true
            & $updateToolsFolderButtonState
        }
    })

    $hayabusaRulesUpdatePollTimer.Add_Tick({
        if ($null -eq $hayabusaRulesUpdateState.Operation) {
            $hayabusaRulesUpdatePollTimer.Stop()
            return
        }

        if (-not $hayabusaRulesUpdateState.Operation.Handle.IsCompleted) {
            $elapsed = [int]((Get-Date) - $hayabusaRulesUpdateState.Operation.Started).TotalSeconds
            $statusLabel.Text = "Updating Hayabusa rules... ${elapsed}s"
            return
        }

        $hayabusaRulesUpdatePollTimer.Stop()
        try {
            $result = $hayabusaRulesUpdateState.Operation.PowerShell.EndInvoke($hayabusaRulesUpdateState.Operation.Handle)
            foreach ($errorRecord in $hayabusaRulesUpdateState.Operation.PowerShell.Streams.Error) {
                $errorMessage = "Hayabusa rules update runspace error: $($errorRecord.Exception.Message)"
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$errorMessage`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $errorMessage
            }

            $resultMessage = "$($result.ModuleId): $($result.Status) - $($result.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$resultMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message $resultMessage
            if (-not [string]::IsNullOrWhiteSpace([string]$result.CommandLine)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Command line: $($result.CommandLine)`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Command line hint: $($result.CommandLine)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$result.WorkingDirectory)) {
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Working directory: $($result.WorkingDirectory)`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Hayabusa rules update working directory: $($result.WorkingDirectory)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$result.StandardOutput)) {
                $standardOutputText = (ConvertTo-IbisGuiDisplayText -Text $result.StandardOutput -StripAnsi).Trim()
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Output:`r`n$standardOutputText`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Hayabusa update-rules stdout: $standardOutputText"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$result.StandardError)) {
                $standardErrorText = (ConvertTo-IbisGuiDisplayText -Text $result.StandardError -StripAnsi).Trim()
                Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Error output:`r`n$standardErrorText`r`n" -StripAnsi
                Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'WARN' -Message "Hayabusa update-rules stderr: $standardErrorText"
            }

            if ($result.Status -eq 'Completed') {
                $statusLabel.Text = 'Hayabusa rules updated'
            }
            else {
                $statusLabel.Text = 'Hayabusa rules update failed'
            }
        }
        catch {
            $updateErrorMessage = "Hayabusa rules update failed: $($_.Exception.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$updateErrorMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $updateErrorMessage
            $statusLabel.Text = 'Hayabusa rules update failed'
        }
        finally {
            Stop-IbisToolInstallRunspace -Operation $hayabusaRulesUpdateState.Operation
            $hayabusaRulesUpdateState.Operation = $null
            $downloadMissingToolsButton.Enabled = $true
            $updateHayabusaRulesButton.Enabled = $true
            $checkToolsButton.Enabled = $true
            $toolGuidanceButton.Enabled = $true
            & $setDefenderControlsForAdmin
            $toolsBrowseButton.Enabled = $true
            $toolsTextBox.Enabled = $true
            & $updateToolsFolderButtonState
        }
    })

    $updateHayabusaRulesButton.Add_Click({
        if ($null -ne $hayabusaRulesUpdateState.Operation) {
            $statusLabel.Text = 'Hayabusa rules update already running'
            return
        }

        $hayabusaStatus = Test-IbisToolStatus -ToolsRoot $toolsTextBox.Text -ToolDefinitions $toolDefinitions |
            Where-Object { $_.Id -eq 'hayabusa' } |
            Select-Object -First 1
        if ($null -eq $hayabusaStatus -or -not $hayabusaStatus.Present) {
            Set-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text 'Hayabusa is not present. Download or install Hayabusa before updating rules.' -StripAnsi
            $statusLabel.Text = 'Hayabusa is missing'
            return
        }

        $message = "Ibis will run Hayabusa update-rules from its install directory:`r`n`r`n$($hayabusaStatus.ExpectedPath)`r`n`r`nThis downloads the latest rules and replaces files in Hayabusa's rules folder. Continue?"
        $choice = [System.Windows.Forms.MessageBox]::Show($message, 'Update Hayabusa Rules', 'YesNo', 'Warning')
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            $statusLabel.Text = 'Hayabusa rules update cancelled'
            return
        }

        $downloadMissingToolsButton.Enabled = $false
        $updateHayabusaRulesButton.Enabled = $false
        $checkToolsButton.Enabled = $false
        $toolGuidanceButton.Enabled = $false
        $checkDefenderExclusionsButton.Enabled = $false
        $addDefenderExclusionsButton.Enabled = $false
        $removeDefenderExclusionsButton.Enabled = $false
        $toolsBrowseButton.Enabled = $false
        $openToolsFolderButton.Enabled = $false
        $toolsTextBox.Enabled = $false
        $toolGuidanceTextBox.Clear()
        Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "Starting Hayabusa rules update...`r`n" -StripAnsi
        Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "The GUI should remain responsive while Hayabusa syncs its rules folder.`r`n" -StripAnsi
        $statusLabel.Text = 'Updating Hayabusa rules'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message "Starting Hayabusa rules update for $($hayabusaStatus.ExpectedPath)."

        try {
            $hayabusaRulesUpdateState.Operation = Start-IbisHayabusaRulesUpdateRunspace -ProjectRoot $ProjectRoot -ToolsRoot $toolsTextBox.Text
            $hayabusaRulesUpdatePollTimer.Start()
        }
        catch {
            $updateStartErrorMessage = "Unable to start Hayabusa rules update: $($_.Exception.Message)"
            Add-IbisTextBoxDisplayText -TextBox $toolGuidanceTextBox -Text "$updateStartErrorMessage`r`n" -StripAnsi
            Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Level 'ERROR' -Message $updateStartErrorMessage
            $statusLabel.Text = 'Hayabusa rules update failed to start'
            if ($null -ne $hayabusaRulesUpdateState.Operation) {
                Stop-IbisToolInstallRunspace -Operation $hayabusaRulesUpdateState.Operation
            }
            $hayabusaRulesUpdateState.Operation = $null
            $downloadMissingToolsButton.Enabled = $true
            $updateHayabusaRulesButton.Enabled = $true
            $checkToolsButton.Enabled = $true
            $toolGuidanceButton.Enabled = $true
            & $setDefenderControlsForAdmin
            $toolsBrowseButton.Enabled = $true
            $toolsTextBox.Enabled = $true
            & $updateToolsFolderButtonState
        }
    })

    $processingPollTimer.Add_Tick({
        if ($null -eq $processingState.Operation) {
            $processingPollTimer.Stop()
            return
        }

        & $appendProcessingProgress

        if (-not $processingState.Operation.Handle.IsCompleted) {
            $elapsed = [int]((Get-Date) - $processingState.Operation.Started).TotalSeconds
            if ([string]::IsNullOrWhiteSpace($runProgressLabel.Text) -or $runProgressLabel.Text -eq 'Processing modules running') {
                $runProgressLabel.Text = "Processing modules running... ${elapsed}s"
                $statusLabel.Text = $runProgressLabel.Text
            }
            return
        }

        $processingPollTimer.Stop()
        $completionNotificationMessage = $null
        $completionNotificationTitle = 'Ibis Processing Complete'
        $completionNotificationIcon = 'Information'
        try {
            & $appendProcessingProgress
            $moduleResults = @($processingState.Operation.PowerShell.EndInvoke($processingState.Operation.Handle))
            foreach ($errorRecord in $processingState.Operation.PowerShell.Streams.Error) {
                Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Processing runspace error: $($errorRecord.Exception.Message)"
            }

            foreach ($moduleResult in $moduleResults) {
                if ($moduleResult.ErrorMessage) {
                    Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "$($moduleResult.ModuleId) failed: $($moduleResult.ErrorMessage)"
                }
                if ($null -ne $moduleResult.Result) {
                    Add-IbisFileOperationHints -Result $moduleResult.Result -LogFilePath $sessionLogPath
                }
                if ($moduleResult.Hostname -and $moduleResult.Hostname -ne 'Unknown') {
                    $hostTextBox.Text = $moduleResult.Hostname
                }
                if ($moduleResult.OutputRoot) {
                    $outputTextBox.Text = $moduleResult.OutputRoot
                }
            }

            if ($processingState.CancelRequested) {
                $statusLabel.Text = 'Processing module run cancelled'
                $runProgressLabel.Text = 'Run cancelled'
                Add-IbisLogLine -LogTextBox $logTextBox -Message 'Processing module run cancelled'
                $completionNotificationMessage = 'Ibis processing was cancelled. The current module was allowed to finish before the run stopped.'
                $completionNotificationTitle = 'Ibis Processing Cancelled'
                $completionNotificationIcon = 'Warning'
            }
            else {
                $statusLabel.Text = 'Processing module run complete'
                $runProgressLabel.Text = 'Run complete'
                Add-IbisLogLine -LogTextBox $logTextBox -Message 'Processing module run complete'
                $completionNotificationMessage = 'Ibis processing run complete.'
            }
            & $updateOutputWarning
        }
        catch {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Processing run failed: $($_.Exception.Message)"
            $statusLabel.Text = 'Processing run failed'
            $runProgressLabel.Text = 'Run failed'
            $completionNotificationMessage = "Ibis processing run failed:`r`n`r`n$($_.Exception.Message)"
            $completionNotificationTitle = 'Ibis Processing Failed'
            $completionNotificationIcon = 'Error'
        }
        finally {
            Stop-IbisToolInstallRunspace -Operation $processingState.Operation
            if (-not [string]::IsNullOrWhiteSpace($processingState.ProgressPath) -and (Test-Path -LiteralPath $processingState.ProgressPath)) {
                try {
                    Remove-Item -LiteralPath $processingState.ProgressPath -Force
                }
                catch {
                }
            }
            $processingState.Operation = $null
            $processingState.ProgressPath = $null
            $processingState.ProgressLineCount = 0
            if (-not [string]::IsNullOrWhiteSpace($processingState.ControlPath) -and (Test-Path -LiteralPath $processingState.ControlPath)) {
                try {
                    Remove-Item -LiteralPath $processingState.ControlPath -Force
                }
                catch {
                }
            }
            $processingState.ControlPath = $null
            $processingState.IsPaused = $false
            $processingState.CancelRequested = $false
            $runProgressBar.MarqueeAnimationSpeed = 0
            $runProgressBar.Value = 0
            $runProgressBar.Visible = $false
            $form.UseWaitCursor = $false
            foreach ($checkBox in $moduleCheckboxes) {
                $checkBox.Enabled = $true
            }
            if ($moduleCheckboxById.ContainsKey('hayabusa') -and $moduleCheckboxById.ContainsKey('takajo')) {
                $moduleCheckboxById['takajo'].Enabled = [bool]$moduleCheckboxById['hayabusa'].Checked
                if (-not $moduleCheckboxById['hayabusa'].Checked) {
                    $moduleCheckboxById['takajo'].Checked = $false
                }
            }
            if ($moduleCheckboxById.ContainsKey('eventlogs') -and $moduleCheckboxById.ContainsKey('duckdb-eventlogs')) {
                $moduleCheckboxById['duckdb-eventlogs'].Enabled = [bool]$moduleCheckboxById['eventlogs'].Checked
                if (-not $moduleCheckboxById['eventlogs'].Checked) {
                    $moduleCheckboxById['duckdb-eventlogs'].Checked = $false
                }
            }
            $runProcessingModulesButton.Enabled = $true
            $pauseProcessingButton.Enabled = $false
            $pauseProcessingButton.Text = 'Pause'
            $cancelProcessingButton.Enabled = $false
            $checkEvidenceButton.Enabled = $true
            $extractHostNameButton.Enabled = $true
            $sourceBrowseButton.Enabled = $true
            $outputBrowseButton.Enabled = $true
            $sourceTextBox.Enabled = $true
            $outputTextBox.Enabled = $true
            $hostTextBox.Enabled = $true
            & $updateOutputWarning
            if (-not [string]::IsNullOrWhiteSpace($completionNotificationMessage)) {
                & $showProcessingCompletionNotification -Message $completionNotificationMessage -Title $completionNotificationTitle -Icon $completionNotificationIcon
            }
        }
    })

    $checkEvidenceButton.Add_Click({
        $result = Test-IbisEvidenceRoot -SourceRoot $sourceTextBox.Text
        Add-IbisLogLine -LogTextBox $logTextBox -Message "Evidence source exists: $($result.Present)"
        Add-IbisLogLine -LogTextBox $logTextBox -Message "Looks like Windows evidence: $($result.LooksLikeWindowsEvidence)"
        $boundaryResult = & $testSourceWriteBoundary
        Add-IbisLogLine -LogTextBox $logTextBox -Message "Source read-only boundary check passed: $($boundaryResult.Passed)"
        foreach ($violation in $boundaryResult.Violations) {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "$($violation.Name) is inside the source: $($violation.Path)"
        }
        foreach ($check in $result.Checks) {
            Add-IbisLogLine -LogTextBox $logTextBox -Message "$($check.RelativePath): $($check.Present)"
        }
        $statusLabel.Text = "Evidence check complete"
    })

    $pauseProcessingButton.Add_Click({
        if ($null -eq $processingState.Operation -or [string]::IsNullOrWhiteSpace($processingState.ControlPath)) {
            return
        }

        try {
            if ($processingState.IsPaused) {
                Set-IbisProcessingControlState -ControlPath $processingState.ControlPath -State 'Running'
                $processingState.IsPaused = $false
                $pauseProcessingButton.Text = 'Pause'
                Add-IbisLogLine -LogTextBox $logTextBox -Message 'Processing resume requested. The next module will start when the worker observes the request.'
                $statusLabel.Text = 'Processing resume requested'
            }
            else {
                Set-IbisProcessingControlState -ControlPath $processingState.ControlPath -State 'Paused'
                $processingState.IsPaused = $true
                $pauseProcessingButton.Text = 'Resume'
                Add-IbisLogLine -LogTextBox $logTextBox -Message 'Processing pause requested. The current module will finish, then Ibis will pause before the next module.'
                $statusLabel.Text = 'Processing pause requested'
            }
        }
        catch {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Unable to update processing pause state: $($_.Exception.Message)"
        }
    })

    $cancelProcessingButton.Add_Click({
        if ($null -eq $processingState.Operation -or [string]::IsNullOrWhiteSpace($processingState.ControlPath)) {
            return
        }

        $choice = [System.Windows.Forms.MessageBox]::Show(
            "Ibis will cancel the processing run before the next module starts. A tool that is already running will be allowed to finish first.`r`n`r`nContinue?",
            'Cancel Processing',
            'YesNo',
            'Warning'
        )
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        try {
            Set-IbisProcessingControlState -ControlPath $processingState.ControlPath -State 'CancelRequested'
            $processingState.CancelRequested = $true
            $pauseProcessingButton.Enabled = $false
            $cancelProcessingButton.Enabled = $false
            Add-IbisLogLine -LogTextBox $logTextBox -Message 'Processing cancel requested. The current module will finish, then Ibis will stop before the next module.'
            $statusLabel.Text = 'Processing cancel requested'
        }
        catch {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Unable to request processing cancellation: $($_.Exception.Message)"
        }
    })

    $extractHostNameButton.Add_Click({
        try {
            if (& $stopIfSourceWriteBoundaryFails -ActionName 'hostname extraction') {
                return
            }

            $result = Invoke-IbisExtractHostName `
                -ToolsRoot $toolsTextBox.Text `
                -ToolDefinitions $toolDefinitions `
                -SourceRoot $sourceTextBox.Text `
                -OutputRoot $outputTextBox.Text `
                -Hostname $hostTextBox.Text

            Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
            if ($result.OutputPath) {
                Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputPath)"
            }
            Add-IbisCommandLineHints -LogTextBox $logTextBox -Result $result -LogFilePath $sessionLogPath
            if ($result.HostName -and $result.HostName -ne 'Unknown') {
                $hostTextBox.Text = $result.HostName
                if ($result.HostOutputRoot) {
                    $outputTextBox.Text = $result.HostOutputRoot
                }
                $statusLabel.Text = "Hostname extracted: $($result.HostName)"
            }
            else {
                $hostTextBox.Text = $config.defaultHostname
                $statusLabel.Text = 'Hostname extraction completed with no hostname found'
            }
            & $updateOutputWarning
        }
        catch {
            $hostTextBox.Text = $config.defaultHostname
            & $updateOutputWarning
            Add-IbisLogLine -LogTextBox $logTextBox -Message "extract-hostname failed: $($_.Exception.Message)"
            $statusLabel.Text = 'Hostname extraction failed'
        }
    })

    $runProcessingModulesButton.Add_Click({
        if ($null -ne $processingState.Operation) {
            $statusLabel.Text = 'Processing modules already running'
            return
        }

        $selectedModules = @()
        foreach ($checkBox in $moduleCheckboxes) {
            if ($checkBox.Checked) {
                $selectedModules += $checkBox.Tag
            }
        }

        $hayabusaSelected = @($selectedModules | Where-Object { $_.id -eq 'hayabusa' }).Count -gt 0
        $takajoSelected = @($selectedModules | Where-Object { $_.id -eq 'takajo' }).Count -gt 0
        $eventLogsSelected = @($selectedModules | Where-Object { $_.id -eq 'eventlogs' }).Count -gt 0
        $duckDbEventLogsSelected = @($selectedModules | Where-Object { $_.id -eq 'duckdb-eventlogs' }).Count -gt 0
        if ($takajoSelected -and -not $hayabusaSelected) {
            Add-IbisLogLine -LogTextBox $logTextBox -Message 'Takajo requires Hayabusa to be selected and run first.'
            $statusLabel.Text = 'Takajo requires Hayabusa'
            return
        }
        if ($duckDbEventLogsSelected -and -not $eventLogsSelected) {
            Add-IbisLogLine -LogTextBox $logTextBox -Message 'DuckDB event log summaries require Windows Event Logs to be selected and run first.'
            $statusLabel.Text = 'DuckDB requires Event Logs'
            return
        }

        $implementedProcessingModules = @($selectedModules | Where-Object { $_.id -eq 'system-summary' -or $_.id -eq 'velociraptor-results' -or $_.id -eq 'registry' -or $_.id -eq 'amcache' -or $_.id -eq 'appcompatcache' -or $_.id -eq 'prefetch' -or $_.id -eq 'ntfs-metadata' -or $_.id -eq 'srum' -or $_.id -eq 'user-artifacts' -or $_.id -eq 'eventlogs' -or $_.id -eq 'duckdb-eventlogs' -or $_.id -eq 'hayabusa' -or $_.id -eq 'takajo' -or $_.id -eq 'chainsaw' -or $_.id -eq 'ual' -or $_.id -eq 'browser-history' -or $_.id -eq 'forensic-webhistory' -or $_.id -eq 'usb' })
        if ($implementedProcessingModules.Count -eq 0) {
            Add-IbisLogLine -LogTextBox $logTextBox -Message 'No implemented processing modules are selected yet.'
            $statusLabel.Text = 'No processing modules selected'
            return
        }

        if (& $stopIfSourceWriteBoundaryFails -ActionName 'processing run') {
            return
        }

        foreach ($checkBox in $moduleCheckboxes) {
            $checkBox.Enabled = $false
        }
        $runProcessingModulesButton.Enabled = $false
        $pauseProcessingButton.Enabled = $true
        $pauseProcessingButton.Text = 'Pause'
        $cancelProcessingButton.Enabled = $true
        $checkEvidenceButton.Enabled = $false
        $extractHostNameButton.Enabled = $false
        $sourceBrowseButton.Enabled = $false
        $outputBrowseButton.Enabled = $false
        $sourceTextBox.Enabled = $false
        $outputTextBox.Enabled = $false
        $hostTextBox.Enabled = $false
        $runProgressBar.Visible = $true
        $runProgressBar.Style = 'Continuous'
        $runProgressBar.Minimum = 0
        $runProgressBar.Maximum = [math]::Max(1, $implementedProcessingModules.Count)
        $runProgressBar.Value = 0
        $runProgressBar.MarqueeAnimationSpeed = 0
        $form.UseWaitCursor = $true

        try {
            Add-IbisLogLine -LogTextBox $logTextBox -Message "Starting selected processing module run: $($implementedProcessingModules.Count) module(s)."
            $runProgressLabel.Text = 'Processing modules running'
            $statusLabel.Text = 'Processing modules running'
            $processingState.ProgressPath = Join-Path ([System.IO.Path]::GetTempPath()) ('IbisProcessingProgress-' + [System.Guid]::NewGuid().ToString() + '.jsonl')
            $processingState.ControlPath = Join-Path ([System.IO.Path]::GetTempPath()) ('IbisProcessingControl-' + [System.Guid]::NewGuid().ToString() + '.json')
            $processingState.ProgressLineCount = 0
            $processingState.IsPaused = $false
            $processingState.CancelRequested = $false
            Set-IbisProcessingControlState -ControlPath $processingState.ControlPath -State 'Running'
            $processingState.Operation = Start-IbisProcessingRunspace `
                -ProjectRoot $ProjectRoot `
                -ToolsRoot $toolsTextBox.Text `
                -ToolDefinitions $toolDefinitions `
                -Modules $implementedProcessingModules `
                -SourceRoot $sourceTextBox.Text `
                -OutputRoot $outputTextBox.Text `
                -Hostname $hostTextBox.Text `
                -ProgressPath $processingState.ProgressPath `
                -ControlPath $processingState.ControlPath
            $processingPollTimer.Start()
        }
        catch {
            Add-IbisLogLine -LogTextBox $logTextBox -Level 'ERROR' -Message "Unable to start background processing: $($_.Exception.Message)"
            if ($null -ne $processingState.Operation) {
                Stop-IbisToolInstallRunspace -Operation $processingState.Operation
            }
            $processingState.Operation = $null
            $processingState.ProgressPath = $null
            $processingState.ProgressLineCount = 0
            if (-not [string]::IsNullOrWhiteSpace($processingState.ControlPath) -and (Test-Path -LiteralPath $processingState.ControlPath)) {
                try {
                    Remove-Item -LiteralPath $processingState.ControlPath -Force
                }
                catch {
                }
            }
            $processingState.ControlPath = $null
            $processingState.IsPaused = $false
            $processingState.CancelRequested = $false
            $runProgressBar.MarqueeAnimationSpeed = 0
            $runProgressBar.Value = 0
            $runProgressBar.Visible = $false
            $form.UseWaitCursor = $false
            foreach ($checkBox in $moduleCheckboxes) {
                $checkBox.Enabled = $true
            }
            & $updateModuleDependencies
            $runProcessingModulesButton.Enabled = $true
            $pauseProcessingButton.Enabled = $false
            $pauseProcessingButton.Text = 'Pause'
            $cancelProcessingButton.Enabled = $false
            $checkEvidenceButton.Enabled = $true
            $extractHostNameButton.Enabled = $true
            $sourceBrowseButton.Enabled = $true
            $outputBrowseButton.Enabled = $true
            $sourceTextBox.Enabled = $true
            $outputTextBox.Enabled = $true
            $hostTextBox.Enabled = $true
            $statusLabel.Text = 'Processing failed to start'
        }

        return

        try {
            Add-IbisLogLine -LogTextBox $logTextBox -Message "Starting selected processing module run: $($implementedProcessingModules.Count) module(s)."
            for ($moduleIndex = 0; $moduleIndex -lt $implementedProcessingModules.Count; $moduleIndex++) {
                $module = $implementedProcessingModules[$moduleIndex]
                $progressText = "Running $($module.name) ($($moduleIndex + 1) of $($implementedProcessingModules.Count))"
                $runProgressLabel.Text = $progressText
                $statusLabel.Text = $progressText
                Add-IbisLogLine -LogTextBox $logTextBox -Message $progressText
                $form.Refresh()
                [System.Windows.Forms.Application]::DoEvents()
                $result = $null
                $fileAuditBefore = Get-IbisFileSystemSnapshot -RootPath $outputTextBox.Text

                if ($module.id -eq 'system-summary') {
                    try {
                        $result = Invoke-IbisSystemSummary `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputPath)"
                        if ($result.HostName -and $result.HostName -ne 'Unknown') {
                            $hostTextBox.Text = $result.HostName
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "system-summary failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'velociraptor-results') {
                    try {
                        $result = Invoke-IbisVelociraptorResultsCopy `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.SourcePath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Source: $($result.SourcePath)"
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputPath)"
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "velociraptor-results failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'registry') {
                    try {
                        $result = Invoke-IbisWindowsRegistryHives `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "registry failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'amcache') {
                    try {
                        $result = Invoke-IbisAmcache `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        }
                        if ($result.JsonPath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "amcache failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'appcompatcache') {
                    try {
                        $result = Invoke-IbisAppCompatCache `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        }
                        if ($result.JsonPath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "appcompatcache failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'prefetch') {
                    try {
                        $result = Invoke-IbisPrefetch `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        }
                        if ($result.JsonPath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "prefetch failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'srum') {
                    try {
                        $result = Invoke-IbisSrum `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        }
                        if ($result.JsonPath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "srum failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'user-artifacts') {
                    try {
                        $result = Invoke-IbisUserArtifacts `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)"
                        }
                        if ($result.JsonPath) {
                            Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)"
                        }
                        if ($result.HostOutputRoot) {
                            $outputTextBox.Text = $result.HostOutputRoot
                        }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "user-artifacts failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'eventlogs') {
                    try {
                        $result = Invoke-IbisEvtxECmdEventLogs `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "eventlogs failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'duckdb-eventlogs') {
                    try {
                        $result = Invoke-IbisDuckDbEventLogSummary `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text `
                            -ProjectRoot $ProjectRoot
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "duckdb-eventlogs failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'hayabusa') {
                    try {
                        $result = Invoke-IbisHayabusaEventLogs `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "hayabusa failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'takajo') {
                    try {
                        $result = Invoke-IbisTakajoEventLogs `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "takajo failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'chainsaw') {
                    try {
                        $result = Invoke-IbisChainsawEventLogs `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "chainsaw failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'ual') {
                    try {
                        $result = Invoke-IbisUserAccessLogsSum `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "ual failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'browser-history') {
                    try {
                        $result = Invoke-IbisBrowsingHistoryView `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "browser-history failed: $($_.Exception.Message)"
                    }
                }
                elseif ($module.id -eq 'usb') {
                    try {
                        $result = Invoke-IbisParseUsbArtifacts `
                            -ToolsRoot $toolsTextBox.Text `
                            -ToolDefinitions $toolDefinitions `
                            -SourceRoot $sourceTextBox.Text `
                            -OutputRoot $outputTextBox.Text `
                            -Hostname $hostTextBox.Text
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "$($result.ModuleId): $($result.Status) - $($result.Message)"
                        if ($result.OutputDirectory) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Output: $($result.OutputDirectory)" }
                        if ($result.JsonPath) { Add-IbisLogLine -LogTextBox $logTextBox -Message "Summary: $($result.JsonPath)" }
                        if ($result.HostOutputRoot) { $outputTextBox.Text = $result.HostOutputRoot }
                    }
                    catch {
                        Add-IbisLogLine -LogTextBox $logTextBox -Message "usb failed: $($_.Exception.Message)"
                    }
                }

                if ($null -ne $result) {
                    Add-IbisCommandLineHints -LogTextBox $logTextBox -Result $result -LogFilePath $sessionLogPath
                    Add-IbisFileOperationHints -Result $result -LogFilePath $sessionLogPath
                }
                $fileAuditAfter = Get-IbisFileSystemSnapshot -RootPath $fileAuditBefore.RootPath
                Add-IbisFileSystemChangeLog -Before $fileAuditBefore -After $fileAuditAfter -Context $module.name -LogFilePath $sessionLogPath
            }

            $statusLabel.Text = 'Processing module run complete'
            $runProgressLabel.Text = 'Run complete'
            & $updateOutputWarning
        }
        finally {
            $runProgressBar.MarqueeAnimationSpeed = 0
            $runProgressBar.Visible = $false
            $form.UseWaitCursor = $false
            foreach ($checkBox in $moduleCheckboxes) {
                $checkBox.Enabled = $true
            }
            if ($moduleCheckboxById.ContainsKey('hayabusa') -and $moduleCheckboxById.ContainsKey('takajo')) {
                $moduleCheckboxById['takajo'].Enabled = [bool]$moduleCheckboxById['hayabusa'].Checked
                if (-not $moduleCheckboxById['hayabusa'].Checked) {
                    $moduleCheckboxById['takajo'].Checked = $false
                }
            }
            if ($moduleCheckboxById.ContainsKey('eventlogs') -and $moduleCheckboxById.ContainsKey('duckdb-eventlogs')) {
                $moduleCheckboxById['duckdb-eventlogs'].Enabled = [bool]$moduleCheckboxById['eventlogs'].Checked
                if (-not $moduleCheckboxById['eventlogs'].Checked) {
                    $moduleCheckboxById['duckdb-eventlogs'].Checked = $false
                }
            }
            $runProcessingModulesButton.Enabled = $true
            $checkEvidenceButton.Enabled = $true
            $extractHostNameButton.Enabled = $true
            $sourceBrowseButton.Enabled = $true
            $outputBrowseButton.Enabled = $true
            $sourceTextBox.Enabled = $true
            $outputTextBox.Enabled = $true
            $hostTextBox.Enabled = $true
            & $updateOutputWarning
        }
    })

    $form.Add_FormClosing({
        & $saveConfigPaths -Reason 'GUI session closing'
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message 'Ibis GUI session closing.'
    })

    $form.Add_FormClosed({
        Write-IbisGuiLogFileLine -LogFilePath $sessionLogPath -Message 'Ibis GUI session closed.'
    })

    $form.Add_Shown({
        $form.Activate()
        & $refreshToolStatusList
    })
    [void]$form.ShowDialog()
}

Export-ModuleMember -Function Show-IbisGui
Export-ModuleMember -Function ConvertTo-IbisGuiDisplayText
Export-ModuleMember -Function Start-IbisHayabusaRulesUpdateRunspace
Export-ModuleMember -Function Start-IbisProcessingRunspace
Export-ModuleMember -Function Stop-IbisToolInstallRunspace
