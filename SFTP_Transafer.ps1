
<#
Purpose- To download the file from SFTP server and to send email
Developer - K.Janarthanan
Date - 13/5/2021
#>

#####################
# Parameter section #
#####################

$Global:LogFile = "$PSScriptRoot\File_Transfer.log" #Log file location
$Global:SourceFile = "user1/sshd_config" # File that needs to be downloaded
$Global:TargetFolder = "E:\Upwork\Upwork_SFTP_Transfer\" # Final destination folder
$Global:LibraryPath = "WinSCPnet.dll" # WinSCP library path
$Global:HostName = "1.52.227.29" #Server IP / Host Name
$Global:UserName = "sf_user01" # SFTP user name
$Global:Password = "Hello" #Password


#Logging function
function Write-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Validateset("INFO","ERR","WARN")]
        [string]$Type="INFO"
    )

    if(-not(Test-Path -path $LogFile.Replace($LogFile.split("\")[-1],"")))
    {
        New-Item -Path $LogFile.Replace($LogFile.split("\")[-1],"") -ItemType "directory" -Force
    }

    $DateTime = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
    $FinalMessage = "[{0}]::[{1}]::[{2}]" -f $DateTime,$Type,$Message

    $FinalMessage | Out-File -FilePath $LogFile -Append
}

#Capture File Transfer Error message
function FileTransferredError
{
    param($File)

    if ($File.Error -eq $Null)
    {
        Write-Log "Transfer of file $($File.FileName) is succeeded"
        return $Null
    }
    else
    {
        Write-Log "Transfer of file $($File.FileName) is failed. Cause is $($File.Error)" -Type ERR
        return $File.Error
    }
}

#Main Function
try 
{
    # Load WinSCP .NET assembly
    Add-Type -Path $LibraryPath

    # Set up session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $HostName
        UserName = $UserName
        Password = $Password
    }

    $sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = "true"
    $session = New-Object WinSCP.Session
    $Flag = $False

    for($i=0; $i -lt 5; $i++)
    {
        try
        {
            # Connect
            Write-Log "Connecting to SFTP Server"
            $session.Open($sessionOptions)

            $transferOptions = New-Object WinSCP.TransferOptions
            $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

            #Rename file
            $TargetFile = "$TargetFolder\{0}" -f ($SourceFile.split("/")[-1]) 
            
            if(Test-Path -Path $TargetFile -PathType Leaf)
            {
                try 
                {
                    Write-Log "Already a file is found at path $TargetFile, therefore going to rename the file"
                    $FinalFileName = "$TargetFolder\{0}_{1}" -f (Get-Date -Format "MM-dd-yyyy_HH_mm_ss"),($SourceFile.split("/")[-1]) 
                    Rename-Item -Path $TargetFile -NewName $FinalFileName -Force -EA Stop
                    Write-Log "File is renamed as $FinalFileName"
                }
                catch 
                {
                    Write-Log "Error occured while renaming the file : $_ , terminating the script" -Type ERR
                    $session.Dispose()
                    exit 1
                }
                
            }
            $session.add_Failed( { FileTransferredError($_) } )
            $transferResult = $session.GetFiles($SourceFile, $TargetFolder, $False, $transferOptions)
            $session.remove_Failed( { FileTransferredError($_) } )

            # Throw on any error
            $transferResult.Check()

            # Final printing of successful file transferred
            foreach ($transfer in $transferResult.Transfers)
            {
                Write-Log "Download of $($transfer.FileName) is succeeded"
            }

            $Flag=$True
        }

        catch
        {
            Write-Log "Error while downloading the file. Summary is : $_, therefore terminating" -Type ERR
            $Flag= $False
        }

        if($Flag)
        {
            break #No any error detected
        }

        Sleep -Seconds 10
    } 

    if($Flag)
    {
        Write-Log "Script done successfully on attempt - $($i+1)"
    }
    else 
    {
        Write-Log "Script terminated with errors after $i re-tries"
    }
 
    $session.Dispose()
    Write-Log "Session Disposed"  
        
    Write-Log "Done with the script"   

}
catch 
{
    Write-Log "Error on main function - $_" -Type ERR
}
