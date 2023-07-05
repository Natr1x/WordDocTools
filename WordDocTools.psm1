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

<#
    .Synopsis
    Extract and join the text from the runs in a '<w:p>' node.

    .Description
    Extract and join the text from the runs in a '<w:p>' node.

    .Parameter Node
    The xml node to extract. Should be a '<w:p>' node.

    .Example
    $xpathargs = Get-XPathArgs -Path '.\MyExample.docx'
    $paragraphs = Select-Xml -XPath '//w:p' @xpathargs |% Node
    Get-ParagraphText -Node $paragraphs
#>
function Get-ParagraphText {
    param(
        [Parameter(
            Position = 0,
            ValueFromPipeLine = $True,
            ValueFromPipeLineByPropertyName = $True)]
        [System.Xml.XmlNode[]]$Node
    )
    process {
        Write-Verbose "Get-ParagraphText: Processing Paragraph"
        foreach ($p in $Node) {
            -join ($p.r.t |% {
                if ($_ -is [string]) {
                    $_
                } elseif ($_.'#text') {
                    $_.'#text'
                }
            })
        }
    }
}
Export-ModuleMember -Function Get-ParagraphText

<#
    .Synopsis
    Extract the text from the cells in a '<w:tbl>' node.

    .Description
    Extract the text from the cells in a '<w:tbl>' node. Use something like
    'Select-Xml -XPath "//w:tbl" @worddocargs |% Node' to acquire the nodes.

    .Parameter Node
    The xml node to extract. Should be a '<w:tbl>' node.

    .Parameter CellDelimiter
    Separator to delimit cells with inside a row.

    .Example
    $xpathargs = Get-XPathArgs -Path '.\MyExample.docx'
    $doctables = Select-Xml -XPath '//w:tbl' @xpathargs |% Node
    Get-TableCells -Node ($doctables[0])
#>
function Get-TableCells {
    param(
        [Parameter(
            Position = 0,
            ValueFromPipeLine = $True,
            ValueFromPipeLineByPropertyName = $True)]
        [System.Xml.XmlNode[]]$Node,
        [Parameter(Mandatory = $False)]
        [string]$CellDelimiter = " | "
    )
    process {
        Write-Verbose "Get-TableCells: Processing TableNode"
        foreach ($t in $Node) {
            foreach ($r in $t.tr) {
                $cols = @()
                foreach ($c in $r.tc) {
                    $cols += ,(Get-ParagraphText $c.p)
                }
                $cols -join "$CellDelimiter"
            }
        }
    }
}
Export-ModuleMember -Function Get-TableCells
