
# PATHS/LOCATIONS

$SrcPath = 'Path' #Path to excel source file
$srcName = Read-Host "Enter file name"
$srcPath = "$SrcPath\$srcName"

# SQL query and csv outputs
$querypath = "Path\Query.sql" # Where the query is saved
$servername = "Server Name"
$outputName = 'OutPut' + (get-date -format ('ddMMyyyy')) 
$path_csv = "Path\$outputName.csv"

#### OPEN THE INPUT FILE ####

$excel = New-Object -ComObject Excel.Application    
$sourceWB = $excel.Workbooks.Open($src)
$sourceWS = $sourceWB.Worksheets.Item('Sheet1')
$sourceWS.Activate()
$rowCount1 = $sourceWS.UsedRange.Rows.Count

##### RUN THE SQL QUERY ####

# copy the ID column from the received excel file
$ID_List =''
foreach ($cell in $sourceWS.Range("A2:C$rowCount1")) {
    if ($cell.value2) {
        $ID_List = $ID_List + $cell.value2 +','
    }
}
$ID_list = $ID_List -replace '\s','' # Remove white space
$ID_List = $ID_List.trimend(',') # Remove the last ','

# paste the list into sql and run it
$content = Get-Content $querypath |ForEach-Object {$_ -replace 'List of ID',$ID_List} 
set-content -value $content -path $querypath
Invoke-Sqlcmd -InputFile $querypath -serverinstance $servername | Export-Csv -path $path_csv -NoTypeInformation

# reset the sql file for future use
$content = Get-Content $querypath |ForEach-Object {$_ -replace $ID_List, 'List of ID'} 
set-content -value $content -path $querypath

###### COPY SQL OUTPUT TO THE SOURCE FILE ###########
$sqlOutput = $excel.Workbooks.Open($path_csv)
$sqlOutputWS = $sqlOutput.Worksheets.Item($outputName) # worksheet name of a csv file is the same with the file name
$rowCount2 = $sqlOutputWS.UsedRange.Rows.Count
$sqlOutputRange = $sqlOutputWS.range("A1:K$rowcount2")
$sqlOutputRange.copy()|Out-Null
$sqlOutputRange_copy = $sourceWB.Worksheets.Item("Sheet2").range("A1:K$rowcount2")
$sqlOutputRange_copy.PasteSpecial(-4163)
$sourceWB.Save()

Set-Clipboard -value 'a' # Avoid the clipboard message box
$excel.Quit()