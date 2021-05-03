function ConvertTo-XLSX([System.Data.DataTable]$dt, [string]$xlsx){

    # make sure epplus is here
    if (-Not (Test-Path "epplus.dll")) {throw "epplus.dll not found"}

    # load epplus
    $null = [Reflection.Assembly]::LoadFile((Get-Item "epplus.dll").FullName)

    # create excel file, add sheet, load data, save, cleanup
    $pkg = New-Object OfficeOpenXml.ExcelPackage $xlsx
    $wks = $pkg.Workbook.Worksheets.Add("Sheet1")
    $rng = $wks.Cells['A1'].LoadFromDataTable([System.Data.DataTable] $dt,$true)
    $rng.AutoFitColumns()
    $tbl = $wks.Tables.Add($rng, "Table1")
    $tbl.TableStyle = 38
    $pkg.Save()
    $pkg.Dispose()
    $pkg = $null

    # return path to file
    (Get-Item $xlsx).FullName

}