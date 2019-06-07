
<#

    - Run As Administrator! 
    - Scripts purpose is to move PST’s from source folder to target folder structure necessary for Commvault PST archive user association”
    - If PST file owner set correctly then script is not necessary
    - Currently specific to only copying PSTs from TransVault migration output from Enterprise Vault.
    - Creates log file if doesn’t exist in same directory as script
    - Creates hashsum for file prior and after copying and compares to determine if mismatch
    - Renames source files after copy to indicate they were copied. Not applicable with -move option
    - source and target directory must exist
    - -source = Source directory where PST’s are located.
    - -target = Target main directory where PST’s are to be moved to
    - -move is an optional parameter that moves rather than copies PST's so faster and saves space.

    Example Usage:
        .\CV-PstCopy.ps1 -source C:\scripts\test\srcdir -target C:\scripts\test\trgtdir
        .\CV-PstCopy.ps1 -source C:\scripts\test\srcdir -target C:\scripts\test\trgtdir -move

#>

#Start Script ----------------

#user input values
param  ( [String] $source, [String] $target, [switch] $move)

#for creating uniqueness 
$timestamp = Get-Date -Format o

#Set log file. Can be changed!
$Logfile = ".\CV-PstCopy.log"

#Simple log writer
Function LogWrite {
   
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring

} #End Function LogWrite

#Add header to log
Function AddHeader {
    
    for ($zz=1; $zz -le 100; $zz++) {
    Add-Content -NoNewline $Logfile -value "-"
    }

    Add-Content $Logfile -value "-" 
    Add-Content  $Logfile -value $timestamp
    
    for ($zz=1; $zz -le 100; $zz++) {
    Add-Content -NoNewline $Logfile -value "-"
    }

    Add-Content $Logfile -value "-" 

} #End Function AddHeader

AddHeader

#Check for -source and -target input values as both required
Function CheckInputs {
    
    If ($source -ne "") {
        $global:path0 = $source
        }

    else {
        Write-Output "Oops! You forgot to specify your source directory!"
        LogWrite "Oops! You forgot to specify your source directory!"
        Exit
        }

    If ($source -ne "") {
        $global:path1 = $target
        }

    else {
        Write-Output "Oops! You forgot to specify your target directory!"
        LogWrite "Oops! You forgot to specify your target directory!"
        Exit
        }

} #End Function CheckInputs

CheckInputs

#Verify paths exist
Function GetMyPaths {


    If (Test-Path $Path0) {
        #The source directory exists. Move along.
        }

    Else {
        #md $Path0
        Write-Output "Oops! This source directory doesn't exist! Please, specify a valid directory."
        LogWrite "Oops! This source directory doesn't exist! Please, specify a valid directory."
        Exit
        }

    If (Test-Path $Path1) {
        #The target directory exists. Move along.
        }

    Else {
        #md $Path1
        Write-Output "Oops! This target directory doesn't exist! Please, specify a valid directory."
        LogWrite "Oops! This target directory doesn't exist! Please, specify a valid directory."
        Exit
        }

} #End Function GetMyPaths

#Copy or Move files into folder structure necessary for Commvault archive
Function CopyFiles {

    #Get All PST Files
    $myarray = Get-ChildItem -Path $path0 -Recurse | where {$_.extension -eq ".pst"}
    $mycount = $myarray.Length
    Write-Output "We found $mycount PST files to copy"
    LogWrite "We found $mycount PST files to copy"

    for ($m=0; $m -lt $myarray.length; $m++) {
	
        $y = $myarray[$m] | Select-Object Name
        $PstFile = $myarray[$m] | Select-Object FullName
        
        
        #Only select files that match regular expression
        If ($y -match "Recipients(\.\w{1,})") {
            
            #Generate hash on original PST file for comparison 
            $myname = $PstFile.FullName
            $orighash = Get-FileHash $myName
            
            #Create User Folders necessary to import PST files into Commvault
            $x = $Matches.0
            $z = $x -replace "Recipients.", ""
            $UserFolder = $path1+"\"+$z
            
            #Make target PST file names. I'm sure this could be done better!!!
            $k = $path1+"\"+$z+"\"+$y.Name
            $g = $k -replace "@{Name=", ""
            
            #Randomize Name to ensure all files get copied
            $CopiedPstFile = $g+"_Commvault_"+"$(get-date -format yyyy-MM-ddTHH-mm-ss-ff)"+".pst"
            
            #Check if user folders exists
            if (Test-Path $UserFolder) {
                
                #Copy or Move PST files
                #If user provides -move switch then move files
                if ($move) {
                    Move-Item $PstFile.FullName -Destination $CopiedPstFile
                    }

                #if no -move switch then copy files and rename source files to indicate they were copied
                else {
                    Copy-Item $PstFile.FullName -Destination $CopiedPstFile -Recurse
                    $ee = $myname+"_CopiedbyCommvault"
                    Rename-Item $PstFile.FullName -NewName $ee
                    }
                
                #Generate hash on copied PST file for comparison 
                $newhash = Get-FileHash $CopiedPstFile
                
                If ($newhash.hash -ne $orighash.hash) {
                    Logwrite "Something went wrong with copy of $PstFile"
                    }

                Else {
                    LogWrite "$orighash copied to $newhash"
                    }
            }
    
            else {
                #make user folders
                md $UserFolder | Out-Null
                
                #Copy or Move PST files
                #If user provides -move switch then move files
                if ($move) {
                    Move-Item $PstFile.FullName -Destination $CopiedPstFile
                    }
                
                #if no -move switch then copy files and rename source files to indicate they were copied
                else {
                    Copy-Item $PstFile.FullName -Destination $CopiedPstFile -Recurse
                    $ee = $myname+"_CopiedbyCommvault"
                    Rename-Item $PstFile.FullName -NewName $ee
                    }
                
                #Generate hash on original PST file for comparison 
                $newhash = Get-FileHash $CopiedPstFile
                
                If ($newhash.hash -ne $orighash.hash) {
                    Logwrite "Something went wrong with copy of $PstFile"
                    Write-Output "Something went wrong with copy of $PstFile"
                    }

                Else {
                    LogWrite "$orighash copied to $newhash"
                    }
                
            }
        }

    }
    
}#End Function CopyFiles

GetMyPaths
CopyFiles

#End Script ----------------







