[CmdletBinding()]
param()

function New-HCellWorkbook
{
    [CmdletBinding()]
    param()

    process
    {
        Write-Verbose "Creating a new HCell application..."
        $hcell = New-Object -ComObject "HCell.Application"

        Write-Verbose "Getting the current active workbook..."
        $workbook = $hcell.ActiveWorkbook

        return $workbook
    }
}

function Add-HCellMacroModule
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        $Workbook,

        [Parameter(Mandatory)]
        $Name,

        [Parameter(Mandatory)]
        $Content
    )

    process
    {
        Write-Verbose "Adding a new module $Name..."

        Write-Verbose "    Adding header to the given code..."
        $Content = "Attribute VB_Name = `"$Name`"`r`n" + $Content

        Write-Verbose "    Adding the macro module to the workbook..."
        $module = $Workbook.ScriptObjects.Add(4, $Name, $Content)

        return
    }
}

function Save-HCellWorkbook
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        $Workbook,

        [Parameter(Mandatory)]
        $Path
    )

    process
    {
        Write-Verbose "Saving the workbook..."
        $Workbook.SaveAs($Path)

        Write-Verbose "Closing the HCell application..."
        $Workbook.Application.Quit()
        return
    }
}


# Enumerate all *.vbs files in the src/ directory
$inputs = Get-ChildItem -Path src -Filter *.vbs

# Create a new HCell instance
$workbook = New-HCellWorkbook

# Add all *.vbs files as macro modules to the workbook
foreach ($file in $inputs)
{
    $name = $file.Name.Replace(".vbs", "")
    $content = (Get-Content $file.FullName) -Join "`r`n"
    Add-HCellMacroModule -Workbook $workbook -Name $name -Content $content
}

# Path of the output file
$output = "raytracing.cell"

# Save the workbook
Save-HCellWorkbook -Workbook $workbook -Path $output
