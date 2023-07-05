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
