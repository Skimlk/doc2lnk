function Edit-LnkDocument {
    Param (
        [string] $document_path,
        [string] $delimiter
    )

    $lnk_path = $document_path + '.lnk'
    $temp_directory = 'C:\Windows\Temp'

    $lnk_content = Get-Content -Path $lnk_path -Raw
    $delimiter_index = $lnk_content.IndexOf($delimiter)
    $delimiter_end_index = $delimiter_index + $delimiter.Length
    $document_content = $lnk_content.Substring($delimiter_end_index)
    $document_temp_path = Join-Path -Path $temp_directory -ChildPath (Split-Path $document_path -Leaf)

    # Write document contents to temp document
    Remove-Item -Path $document_temp_path -Force -ErrorAction SilentlyContinue
    Add-Content -Path $document_temp_path -Value $document_content
    Start-Process -FilePath $document_temp_path -Wait

    # Save temp document to end of .lnk file
    $saved_lnk_content = $lnk_content.Substring(0, $delimiter_end_index) + (Get-Content -Path $document_temp_path -Raw)
    Remove-Item -Path $lnk_path -Force -ErrorAction SilentlyContinue
    Add-Content -Path $lnk_path -Value $saved_lnk_content
}