function Get-ExcelCellInfo {
    
    #Load the required functions
    . .\Manage-Functions.ps1
    . Manage-Functions

    # Set-PSRepository -name  "PSGallery" -InstallationPolicy Trusted
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
    # Install-Module ImportExcel -Scope CurrentUser -Confirm:$false #https://github.com/dfinke/ImportExcel


    #Get all excel files from the repo
    $AllFiles = Get-ChildItem -Path ".\DataFiles\*.xlsx"
    

    foreach($file in $AllFiles){
        
        Import-Excel -Path ".\DataFiles\$($file.name)"

        "------------------------------------------------------"
    }




























# # "test" | Out-File "./DataFiles/Test.txt" -Verbose

    # UnZip-File "./DataFiles/Test.docx" "./DataFiles/Temp"

    # Remove-Item "./DataFiles/Ninino.docx" -Force -ErrorAction SilentlyContinue
    # Zip-File "./DataFiles/Ninino.docx" "./DataFiles/Temp"






    # $Myexcel = New-Object -ComObject excel.application
    # $Myexcel.visible = $false
    # $Myworkbook = $Myexcel.workbooks.add()
    # $Sheet1 = $Myworkbook.worksheets.item(1)
    # $Sheet1.name = "Power level"
    # $Sheet1.cells.item(1,1) = 'NAME'
    # $Sheet1.cells.item(1,2) = 'POWER'
    # $Sheet1.cells.item(1,3) = 'TRANSFORMATION'
    # $Sheet1.Range("A1:C1").font.size = 18
    # $Sheet1.Range("A1:C1").font.bold = $true
    # $Sheet1.Range("A1:C1").font.ColorIndex = 2
    # $Sheet1.Range("A1:C1").interior.colorindex = 1

    # $Sheet1.cells.item(2,1) = 'Goku'
    # $Sheet1.cells.item(2,2) = '9500'
    # $Sheet1.cells.item(2,3) = 'SSJ3'

    # $Sheet1.cells.item(3,1) = 'Vegeta'
    # $Sheet1.cells.item(3,2) = '10000'
    # $Sheet1.cells.item(3,3) = 'SSJ2'

    # $Sheet1.cells.item(4,1) = 'Gohan'
    # $Sheet1.cells.item(4,2) = '5000'
    # $Sheet1.cells.item(4,3) = 'SSJ2'

    # $Sheet1.Range("A1:C4").HorizontalAlignment = -4108
    # $Sheet1.Range("A1:C4").VerticalAlignment = -4108

    # $Sheet1.Range("A1:C4").Borders.LineStyle = 1
    # $Sheet1.Columns.AutoFit()
    # $Myfile = 'E:\09-P_Appro\PS-Eval\DataFiles\example.xlsx'
    # $Myexcel.displayalerts = $false
    # $Myworkbook.Saveas($Myfile)
    # $Myexcel.displayalerts = $true




    #Unloading the functions
    . Manage-Functions -remove
}