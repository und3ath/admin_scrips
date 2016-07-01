[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization") 
$chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
$chart1.Width = 3192
$chart1.Height = 800
$chart1.BackColor = [System.Drawing.Color]::White   
[void]$chart1.Titles.Add("ActiveSync Devices per type") 
$chart1.Titles[0].Font = "Arial,13pt" 
$chart1.Titles[0].Alignment = "topLeft" 

$chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
$chartarea.Name = "ChartArea1" 
$chartarea.AxisY.Title = "Number of devices" 
$chartarea.AxisX.Title = "Devices types" 
$chartarea.AxisY.Interval = 1 
$chartarea.AxisX.Interval = 1 
$chart1.ChartAreas.Add($chartarea) 
 
$legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend 
$legend.name = "Legend1" 
$chart1.Legends.Add($legend)  

$datasource = Get-ActiveSyncDevice -ResultSize unlimited | select DeviceType | Group-Object DeviceType

[void]$chart1.Series.Add("DevicesAS") 
$chart1.Series["DevicesAS"].ChartType = "Column" 
$chart1.Series["DevicesAS"].BorderWidth = 3 
$chart1.Series["DevicesAS"].IsVisibleInLegend = $true 
$chart1.Series["DevicesAS"].chartarea = "ChartArea1" 
$chart1.Series["DevicesAS"].Legend = "Legend1" 
$chart1.Series["DevicesAS"].color = "#62B5CC" 
$datasource | ForEach-Object {$chart1.Series["DevicesAS"].Points.addxy( $_.name , $_.count) } 


$dlg = New-Object System.Windows.Forms.SaveFileDialog
$dlg.Filter = "Png File (*.png)|*.png"
$dlg.SupportMultiDottedExtensions = $true
$dlg.InitialDirectory = $PSScriptRoot
$dlg.CheckFileExists = $false

if($dlg.ShowDialog() -eq 'Ok')
{
    $chart1.SaveImage($dlg.FileName,"png")  
}

