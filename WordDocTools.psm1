<#
    .Synopsis
    Extracts 'word/document.xml' and its namespace dictionary from a word document.

    .Description
    Extracts 'word/document.xml' and its namespace dictionary from a word document.
    The returned hashmap can be used as arguments whith powershells 'Select-Xml'.

    .Parameter Path
    Path to the word document to extract.

    .Example
    $xpathargs = Get-XPathArgs -Path '.\MyExample.docx'
    $doctables = Select-Xml -XPath '//w:tbl' @xpathargs |% Node
#>
function Get-XPathArgs {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    $Path = (Resolve-Path $Path).ProviderPath
    try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
            $entry = $zip.GetEntry('word/document.xml')
            try {
                    $reader = [System.IO.StreamReader]::new($entry.Open())
                    [xml]$doc = $reader.ReadToEnd()
            } catch {
                    Write-Error "InnerError: $_"
                    throw
            } finally {
                    ${reader}?.Dispose()
            }
    } catch {
            Write-Error "OuterError: $_"
            throw
    } finally {
            ${zip}?.Dispose()
    }
    if (-not $doc) {
            Write-Error "Doc is null?? $doc"
            return
    }

    $ns = @{}
    $doc.document.Attributes |? Name -like 'xmlns:*'
    |% {$ns.Add("$($_.Name.Substring(6))", $_.Value)}

    $selArg = @{
            Namespace = $ns
            Xml = $doc
    }
    return $selArg
}
Export-ModuleMember -Function Get-XPathArgs

class Delimiters {
    [string]$Run = $null
    [string]$Paragraph = $null
    [string]$Cell = $null
    [string]$Row = $null
}

class WordObject {
    [string] ContentAsText([Delimiters]$delims) {
        return ""
    }
}

class WordRun : WordObject {
    [string[]]$Text

    [string] ContentAsText([Delimiters]$delims) {
        return -join $this.Text
    }
}

class WordParagraph : WordObject {
    [WordRun[]]$Runs

    [string] ContentAsText([Delimiters]$delims) {
        [string[]]$childTexts = $this.Runs |% ContentAsText $delims
        [string]$delim = $delims.Run

        if ($delim -ne $null) {
            return $childTexts -join $delim
        } else {
            return -join $childTexts
        }
    }
}

class WordTableCell : WordObject {
    [WordParagraph[]]$Paragraphs

    [string] ContentAsText([Delimiters]$delims) {
        [string[]]$childTexts = $this.Paragraphs |% ContentAsText $delims
        [string]$delim = $delims.Paragraph

        if ($delim -ne $null) {
            return $childTexts -join $delim
        } else {
            return -join $childTexts
        }
    }
}

class WordTableRow : WordObject {
    [WordTableCell[]]$Cells

    [string] ContentAsText([Delimiters]$delims) {
        [string[]]$childTexts = $this.Cells |% ContentAsText $delims
        [string]$delim = $delims.Cell

        if ($delim -ne $null) {
            return $childTexts -join $delim
        } else {
            return -join $childTexts
        }
    }
}

class WordTable : WordObject {
    [WordTableRow[]]$Rows

    [string] ContentAsText([Delimiters]$delims) {
        [string[]]$childTexts = $this.Rows |% ContentAsText $delims
        [string]$delim = $delims.Row

        if ($delim -ne $null) {
            return $childTexts -join $delim
        } else {
            return -join $childTexts
        }
    }
}

<#
    .Synopsis
    Extract the text from an xml node form a word document.

    .Description
    Extract the text from an xml node form a word document. Use something like
    'Select-Xml -XPath "//w:tbl" @worddocargs |% Node' to acquire the nodes.

    .Parameter Node
    The xml node to extract.

    .Example
    $xpathargs = Get-XPathArgs -Path '.\MyExample.docx'
    $doctables = Select-Xml -XPath '//w:tbl' @xpathargs |% Node
    Get-DocumentItem -Node ($doctables[0])
#>
function Get-DocumentItem {
    param(
        [Parameter(
            Position = 0,
            ValueFromPipeLine = $True,
            ValueFromPipeLineByPropertyName = $True)]
        [System.Xml.XmlNode[]]$Node
    )
    process {
        foreach ($n in $Node) {
            Write-Verbose "Get-DocumentItem: Parsing $($n.Name)"
            switch ($n.Name) {
                'w:tbl' {
                    [WordTable]@{
                        Rows = @(Get-DocumentItem -Node $n.tr @PSBoundParameters)
                    }
                }
                'w:tr' {
                    [WordTableRow]@{
                        Cells = @(Get-DocumentItem -Node $n.tc @PSBoundParameters)
                    }
                }
                'w:tc' {
                    [WordTableCell]@{
                        Paragraphs = @(Get-DocumentItem -Node $n.p @PSBoundParameters)
                    }
                }
                'w:p' {
                    [WordParagraph]@{
                        Runs = @(Get-DocumentItem -Node $n.r @PSBoundParameters)
                    }
                }
                'w:r' {
                    [WordRun]@{
                        Text = @($n.t |% {
                            if ($_ -is [string]) {
                                $_
                            } elseif ($_.'#text') {
                                $_.'#text'
                            }
                        })
                    }
                }
            }
        }
    }
}
Export-ModuleMember -Function Get-DocumentItem

<#
    .Synopsis
    Extract the text content from a document item.

    .Description
    Extract the text content from a document item. Use 'Get-DocumentItem' to
    get items that text can be extracted from.

    .Parameter DocumentItem
    The part of the document to get as text

    .Parameter RunDelimiter
    Delimeter to use between runs

    .Parameter ParagraphDelimiter
    Delimeter to use between paragraphs

    .Parameter CellDelimiter
    Delimeter to use between cells

    .Parameter RowDelimiter
    Delimeter to use between rows

    .Example
    $xpathargs = Get-XPathArgs -Path '.\MyExample.docx'
    $doctables = Select-Xml -XPath '//w:tbl' @xpathargs
    $doctables |Get-DocumentItem |Get-TextFromDocumentItem
#>
function Get-TextFromDocumentItem {
    param(
        [Parameter(
            Position = 0,
            ValueFromPipeLine = $True,
            ValueFromPipeLineByPropertyName = $True)]
        [WordObject[]]$DocumentItem,

        [Parameter(Mandatory = $False)]
        [string]$RunDelimiter = "",

        [Parameter(Mandatory = $False)]
        [string]$ParagraphDelimiter = "",

        [Parameter(Mandatory = $False)]
        [string]$CellDelimiter = " | ",

        [Parameter(Mandatory = $False)]
        [string]$RowDelimiter = [System.Environment]::NewLine
    )
    begin {
        Write-Verbose "Get-TextFromDocumentItem: Begin"
        $delims = [Delimiters]@{
            Run = $RunDelimiter
            Paragraph = $ParagraphDelimiter
            Cell = $CellDelimiter
            Row = $RowDelimiter
        }
    }
    process {
        Write-Verbose "Get-TextFromDocumentItem: Processing DocumentItem"
        foreach ($b in $DocumentItem) {
            $b.ContentAsText($delims)
        }
    }
}
Export-ModuleMember -Function Get-TextFromDocumentItem
