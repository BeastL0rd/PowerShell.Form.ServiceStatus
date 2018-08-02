#  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process 

#Execution Policy Type:
#
#    Restricted - (default) No Script either local, remote or downloaded can be executed on the system.
#    AllSigned - All script that are ran require to be digitally signed.
#    RemoteSigned - All remote scripts (UNC) or downloaded need to be signed.
#    Unrestricted - No signature for any type of script is required.
#


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

Clear-Host

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '1300,600'
$Form.text                       = "Femida Remote"
$Form.TopMost                    = $false

$Button_Close                         = New-Object system.Windows.Forms.Button
$Button_Close.text                    = "Close"
$Button_Close.width                   = 60
$Button_Close.height                  = 30
$Button_Close.location                = New-Object System.Drawing.Point(10,10)
$Button_Close.Font                    = 'Microsoft Sans Serif,10'

$Label_user                      = New-Object system.Windows.Forms.Label
$Label_user.text                 = "User"
$Label_user.AutoSize             = $true
$Label_user.width                = 55
$Label_user.height               = 10
$Label_user.location             = New-Object System.Drawing.Point(290,15)
$Label_user.Font                 = 'Microsoft Sans Serif,10'

$DataGridView1                       = New-Object system.Windows.Forms.DataGridView
$DataGridView1.width                 = 1280
$DataGridView1.height                = 500
$DataGridView1.location              = New-Object System.Drawing.Point(10,80)
$DataGridView1.AllowUserToAddRows    = $false
$DataGridView1.AllowUserToDeleteRows = $false

$Button_get_ip_all                         = New-Object system.Windows.Forms.Button
$Button_get_ip_all.text                    = "Get IP (all)"
$Button_get_ip_all.width                   = 100
$Button_get_ip_all.height                  = 30
$Button_get_ip_all.location                = New-Object System.Drawing.Point(10,45)
$Button_get_ip_all.Font                    = 'Microsoft Sans Serif,10'

$Button_update_status                         = New-Object system.Windows.Forms.Button
$Button_update_status.text                    = "Update Status (all)"
$Button_update_status.width                   = 140
$Button_update_status.height                  = 30
$Button_update_status.location                = New-Object System.Drawing.Point(120,45)
$Button_update_status.Font                    = 'Microsoft Sans Serif,10'

$Label_service                   = New-Object system.Windows.Forms.Label
$Label_service.text              = "Service:"
$Label_service.AutoSize          = $true
$Label_service.width             = 25
$Label_service.height            = 10
$Label_service.location          = New-Object System.Drawing.Point(270,50)
$Label_service.Font              = 'Microsoft Sans Serif,10'

$ComboBox_service                = New-Object system.Windows.Forms.ComboBox
$ComboBox_service.text           = "SRSREGISTRATOR"
$ComboBox_service.width          = 150
$ComboBox_service.height         = 20
$ComboBox_service.location       = New-Object System.Drawing.Point(320,47)
$ComboBox_service.Font           = 'Microsoft Sans Serif,10'

$ProgressBar1                    = New-Object system.Windows.Forms.ProgressBar
$ProgressBar1.width              = 200
$ProgressBar1.height             = 18
$ProgressBar1.location           = New-Object System.Drawing.Point(80,16)

$Form.controls.AddRange(@($Button_Close,$Label_user,$DataGridView1,$Button_get_ip_all,$ProgressBar1,$Button_update_status, $ComboBox_service, $Label_service))

#region gui events {
$DataGridView1.Add_CellClick({ Grid_Cell_Click })
$Button_Close.Add_Click({ Button_Click_Close })
$Button_get_ip_all.Add_Click({ Button_Click_get_ip_all })
$Button_update_status.Add_Click({ Button_Click_update_status })

$ComboBox_service.Add_SelectedIndexChanged({ Item_pick })

#Add Form event 
$Form.add_Load({OnLoadForm_Update}) 

#endregion events }

#endregion GUI }


#Write your logic code here
#Set-Variable -Name "debuger_console" -Value true
#$Global:debuger_console = $false
$Global:debuger_console = $true
$Global:user_a = "administrator"
$Global:current_user 
$Global:credential

$Global:get_process = $false
$Global:progress = 0

$Global:RunspacePoolG = [RunspaceFactory ]::CreateRunspacePool(1, 8)

        $Global:service_name = "Print Spooler"
        




function create_data_grid
{
    $DataGridView1.Columns.Add("pc_name", "PC Name")  
    $DataGridView1.Columns.Insert(1, (New-Object System.Windows.Forms.DataGridViewButtonColumn))    
    $DataGridView1.Columns.Add("pc_ip", "IP")
    $DataGridView1.Columns.Add("service_status", "Service Status")
    $DataGridView1.Columns.Insert(3, (New-Object System.Windows.Forms.DataGridViewButtonColumn))



    $DataGridView1.Columns[1].Name="Get IP"
    $DataGridView1.Columns[3].Name="Get Status"
   

    $DataGridView1.Columns[0].ReadOnly = $true
    $DataGridView1.Columns[2].ReadOnly = $true
    $DataGridView1.Columns[4].ReadOnly = $true

}

function read_text_file
{
    $DataGridView1.Rows.Add("PC68-K259-DIV","Get IP","???","Get","?")
    $DataGridView1.Rows.Add("PC68-K259-1G17","Get IP","???","Get","?")
    $DataGridView1.Rows.Add("PC68-K259-MSUB","Get IP","???","Get","?")
    $DataGridView1.Rows.Add("PC68-S201","Get IP","???","Get","?")
    $DataGridView1.Rows.Add("zzzz","Get IP","???","Get","?")
    

}


function OnLoadForm_Update 
{ 
   

    $Global:current_user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    Log($Global:current_user)

    $Label_user.text                 = "Current user: " + $Global:current_user

   
    $ComboBox_service.Items.add($Global:service_name)
    $ComboBox_service.Items.add("Print Spooler")


    create_data_grid
    read_text_file
    
    $Global:RunspacePoolG.Open()


 $form.refresh() 
} 

function Item_pick{
    
    $Global:service_name = $ComboBox_service.SelectedItem.ToString()

    Log("Global:service_name= " + $Global:service_name)
    Log("ComboBox_service= " + $ComboBox_service.SelectedItem.ToString())

}



function Button_Click_Close { 
 $Form.Close()
} 

function Button_Click_get_ip_all { 

    $Button_get_ip_all.Enabled = $false

    $Global:progress = 0
    $ProgressBar1.Value = $Global:progress
    
    $line_count = $DataGridView1.RowCount

    $Global:progress_max = $line_count


    $RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, 8)   

    $Jobs = @()

    $RunspacePool.Open()
    

    For ($i=0; $i -le $line_count-1; $i++) {

        $Job = [powershell]::Create().AddScript($ScriptBlock_get_ip_address).AddArgument($DataGridView1.Rows[$i].Cells[0].Value) 
        $Job.RunspacePool = $RunspacePool

        $Jobs += New-Object PSObject -Property @{
          RunNum = $_
          Pipe = $Job
          rowIndex = $i      
          Result = $Job.BeginInvoke()
         }
        log("BeginInvoke()")        

        $Global:progress =  [math]::Round($i * 100/$Global:progress_max*2/3)

        $ProgressBar1.Value = $Global:progress
        $ProgressBar1.refresh()

    }


     Do{
        if($ProgressBar1.Value -le 90){
            $Global:progress = $Global:progress + 1
            $ProgressBar1.Value = $Global:progress
            $ProgressBar1.refresh() 
        }
        log($Global:progress)
        Start-Sleep -m 200
    } While ($Jobs.Result.IsCompleted -contains $false)

    $rowIndex = 1
    [string]$ip_address = ""
    ForEach ($Job in $Jobs){   
        $ip_address = $Job.Pipe.EndInvoke($Job.Result)
        $rowIndex = $Job.rowIndex
        log(".")

        if(($ip_address -eq "") -Or ($ip_address -eq $null)){
            $DataGridView1.Rows[$rowIndex].Cells[2].Value = ""
            $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor = "Yellow"
        }else{
            $DataGridView1.Rows[$rowIndex].Cells[2].Value = $ip_address
            $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor = "White"         

        }
    }

    

    
    
    $Jobs.Clear()
    $RunspacePool.Close()

    log("RunspacePool.Close()")
    log($ip_address)



    $ProgressBar1.Value = 100

    $Button_get_ip_all.Enabled = $true

} 


function Button_Click_update_status { 

    $Button_update_status.Enabled = $false
    

    #####################################################

    $Global:progress = 0
    $ProgressBar1.Value = $Global:progress
    
    $line_count = $DataGridView1.RowCount

    $Global:progress_max = $line_count
    [float]$progressf = 0


    #$RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, 8)   

    $Jobs = @()

    #$RunspacePool.Open()
    

    For ($i=0; $i -le $line_count-1; $i++) {

        $Job = [powershell]::Create().AddScript($ScriptBlock_get_service_status)
        $Job.AddArgument($DataGridView1.Rows[$i].Cells[0].Value) 
        $Job.AddArgument($Global:service_name)
        $Job.RunspacePool = $Global:RunspacePoolG

        $Jobs += New-Object PSObject -Property @{
          RunNum = $_
          Pipe = $Job
          rowIndex = $i
          row_updated = $false    
          pcName = $DataGridView1.Rows[$i].Cells[0].Value 
          Result = $Job.BeginInvoke()
         }
        log("BeginInvoke()")        

        $Global:progress =  [math]::Round($i * 100/$Global:progress_max*2/3)

        $ProgressBar1.Value = $Global:progress
        $ProgressBar1.refresh()

    }


     Do{
        if($ProgressBar1.Value -le 90){
            $progressf =  $progressf +0.2
            $Global:progress = $progressf 
            $ProgressBar1.Value = $Global:progress
            $ProgressBar1.refresh() 
        }
        log($Global:progress)
        [string]$service = ""
        ForEach ($Job in $Jobs){   
            if(($Job.Result.IsCompleted -eq $true) -and ($Job.row_updated -eq $false)){
                $service = $Job.Pipe.EndInvoke($Job.Result)
                $rowIndex = $Job.rowIndex
                log("rowIndex = "+ $rowIndex + " " +$service)
                log("Job.Result "+$Job.Result)
                log("pc_name =  "+$Job.pcName)
 

                if(($service -eq "") -Or ($service -eq $null)){
                    $DataGridView1.Rows[$rowIndex].Cells[4].Value = "not found"
                    $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Pink"        
            
                }else {
                    $DataGridView1.Rows[$rowIndex].Cells[4].Value =  $service
                    if($DataGridView1.Rows[$rowIndex].Cells[4].Value -eq "Running"){
                        $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "White" 
                    } else{
                        $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Yellow" 
                    }
                }
                $Job.row_updated = $true
                $Job.Pipe.Dispose()

            }
        }
        
        Start-Sleep -m 200
    } While ($Jobs.Result.IsCompleted -contains $false)

    
        [string]$service = ""
        ForEach ($Job in $Jobs){   
            if(($Job.Result.IsCompleted -eq $true) -and ($Job.row_updated -eq $false)){
                $service = $Job.Pipe.EndInvoke($Job.Result)
                $rowIndex = $Job.rowIndex
                log("rowIndex = "+ $rowIndex + " " +$service)
                log("Result= "+$Job.Result)
                log("pc_name =  "+$Job.pcName)


                if(($service -eq "") -Or ($service -eq $null)){
                    $DataGridView1.Rows[$rowIndex].Cells[4].Value = "not found"
                    $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Pink"        
            
                }else {
                    $DataGridView1.Rows[$rowIndex].Cells[4].Value =  $service
                    if($DataGridView1.Rows[$rowIndex].Cells[4].Value -eq "Running"){
                        $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "White" 
                    } else{
                        $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Yellow" 
                    }
                }
                $Job.row_updated = $true
                $Job.Pipe.Dispose()

            }
        }

    

    
    
    $Jobs.Clear()
    #$RunspacePool.Close()
    #$RunspacePool.Dispose()

    log("RunspacePool.Close()")
    log($service)



    $ProgressBar1.Value = 100


    $Button_update_status.Enabled = $true
} 

function Grid_Cell_Click 
{ 
    

     $rowIndex = $DataGridView1.CurrentRow.Index
     $columnIndex = $DataGridView1.CurrentCell.ColumnIndex

     if ($columnIndex -eq 1){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        get_ip_address($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 2){
       
        $DataGridView1.Rows[$rowIndex].Cells[2].Value = "?"
        $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor = "White"
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 3){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        get_service_status($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 6){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        restart_service($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 7){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        start_service($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 8){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        stop_service($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 

     } elseif($columnIndex -eq 9){
        $Global:progress = 0
        $ProgressBar1.Value = 5
        stop_process($rowIndex,$columnIndex)
        $ProgressBar1.Value = 100
        $ProgressBar1.refresh() 
     }

     log("(cell click) rowIndex ="+ $rowIndex+ " columnIndex=" + $columnIndex)
 
    
} 

$ScriptBlock_get_ip_address = {
   Param (
      [string]$pc_name

   )
   $ip_address = Resolve-DnsName -Type A $pc_name
 
    $Result = ""
    if($ip_address){
        $Result = $ip_address.IPAddress.ToString()
    }else{
        $Result = ""
    }
   
   
   Return $Result
}

function get_ip_address($rowIndex,$columnIndex){

    $RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, 8)
    $RunspacePool.Open()
   
    $Jobs = @()
   
    $Job = [powershell]::Create().AddScript($ScriptBlock_get_ip_address).AddArgument($DataGridView1.Rows[$rowIndex].Cells[0].Value) 
    $Job.RunspacePool = $RunspacePool
 
    $Jobs += New-Object PSObject -Property @{
      RunNum = $_
      Pipe = $Job
      Result = $Job.BeginInvoke()
    }
    log("BeginInvoke()")
    Do{
        if($ProgressBar1.Value -le 90){
            $Global:progress = $Global:progress +1
            $ProgressBar1.Value = $Global:progress
            $ProgressBar1.refresh() 
        }
        log($Global:progress)
        Start-Sleep -m 200
    } While ($Jobs.Result.IsCompleted -contains $false)

    [string]$ip_address = ""
    ForEach ($Job in $Jobs){   
        $ip_address += $Job.Pipe.EndInvoke($Job.Result)

        log(".")
    }

    

    if(($ip_address -eq "") -Or ($ip_address -eq $null)){
        $DataGridView1.Rows[$rowIndex].Cells[2].Value = ""
        $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor = "Yellow"
    }else{
        $DataGridView1.Rows[$rowIndex].Cells[2].Value = $ip_address
        $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor = "White" 
    }
    
    $Jobs.Clear()
    $RunspacePool.Close()
    $RunspacePool.Dispose()

    log("RunspacePool.Close()")
    log($ip_address)


}

$ScriptBlock_get_service_status = {
   Param (
      [string]$pc_name,
      [string]$service_name

   )
   $service = Get-Service -DisplayName $service_name -ComputerName $pc_name -ErrorAction SilentlyContinue
 
    $Result = ""
    if($service){
            $Result =  $service.Status.ToString()                    
        }else {
            $Result = ""
        }
   
   
   Return $Result
}

function get_service_status($rowIndex,$columnIndex){

    if($DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor -eq "Yellow"){
        log("Yellow = " + $DataGridView1.Rows[$rowIndex].Cells[2].Style.backcolor)
    } else {
    
        $RunspacePool = [RunspaceFactory ]::CreateRunspacePool(1, 8)
        $RunspacePool.Open()

        $Jobs = @()
        
   
        $Job = [powershell]::Create().AddScript($ScriptBlock_get_service_status)
        $Job.AddArgument($DataGridView1.Rows[$rowIndex].Cells[0].Value) 
        $Job.AddArgument($Global:service_name) 

        $Job.RunspacePool = $RunspacePool
 
        $Jobs += New-Object PSObject -Property @{
          RunNum = $_
          Pipe = $Job
          Result = $Job.BeginInvoke()
        }
        log("BeginInvoke()")




        Do{
        if($ProgressBar1.Value -le 90){
            $Global:progress = $Global:progress +1
            $ProgressBar1.Value = $Global:progress
            $ProgressBar1.refresh() 
        }
        log($Global:progress)
        Start-Sleep -m 200
        } While ($Jobs.Result.IsCompleted -contains $false)

        [string]$service = ""
        ForEach ($Job in $Jobs){   
            $service += $Job.Pipe.EndInvoke($Job.Result)

            log(".")
        }


        if(($service -eq "") -Or ($service -eq $null)){
            $DataGridView1.Rows[$rowIndex].Cells[4].Value = "not found"
            $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Pink"        
            
        }else {
            $DataGridView1.Rows[$rowIndex].Cells[4].Value =  $service
            if($DataGridView1.Rows[$rowIndex].Cells[4].Value -eq "Running"){
                $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "White" 
            } else{
                $DataGridView1.Rows[$rowIndex].Cells[4].Style.backcolor = "Yellow" 
            }
        }






        $Jobs.Clear()
        $RunspacePool.Close()
        $RunspacePool.Dispose()

        log("RunspacePool.Close()")
        log($service)



             
   

    }
}


function log($text){
    if($Global:debuger_console){
        Write-Host "> " + $text
    }
}


[void]$Form.ShowDialog()
