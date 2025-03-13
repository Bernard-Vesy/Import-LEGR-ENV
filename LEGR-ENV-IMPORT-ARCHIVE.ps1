# - Import le set de donnée dans une variable $DATA

#$data = Import-Excel -Path "\\lehu-fs-1-p\File_Share_Root\Docu\5000 - Vállalati Dokumentumok_Corporate Documents\10 - PDF-ek Share Pointra\RKFT Document list with tags.xlsx" 
$data = Import-Excel -Path "C:\Users\bve\OneDrive - LEMO SA\Applications\LE_Environnement\LE_Environnement.xlsx" -WorksheetName "Archive"


#Migration des fichiers d'un serveur de fichiers vers un site SharePoint
#Connect to PNP sadfasdfasdf
$weburl = "https://lemo.sharepoint.com/sites/LEGR-ENV"
#Connect-PnPOnline -Url $weburl -UseWebLogin
Connect-PnPOnline -Url $weburl -Interactive -ClientId '78d71e5d-4290-4f85-b915-d9958bb940bf'

#Connect-PnPOnline -Url $weburl 

$DocLib = "ARCHIVES"

# - Delete the library content ----------------------------------
#Get-PnPList -Identity $DocLib | Get-PnPListItem -PageSize 100 -ScriptBlock {
#    Param($items) Invoke-PnPQuery } | ForEach-Object { $_.Recycle() | Out-Null
#}

foreach($line in $data)
{
#    $line.'File Name '
#    $line.'File Size'
#$line.'reference'

    $Reference = "" 
    #search file name in the sub directory
    
    $Folder = $line.'FullPath' #"\\ntlemo-webfs-1-p\Portail\" + $line.'reference'
    #$Folder = "\\DCLEMO\LE_Environnement\Conformite Fournisseurs - Copie\AA-NEW -MATERIALS SDS PLASTIC\POM\" + $line.'File Name '
    $NumberOfFiles = 0
    if (Test-Path -Path $Folder) {
        # Path exists!
        $files=Get-ChildItem $Folder
        $NumberOfFiles = $files.Count
        if($NumberOfFiles -eq 1)
        {
            $Reference = $Reference = $files.FullName
        }
        else {
              # write-host " more than 1 file : Count = "  $NumberOfFiles " : " $Folder
        }
    } 
    else {
        write-host "File not exist " $line.'Name'
    }
    
     
    if (![string]::IsNullOrEmpty($Reference))
    {
        $Author = "vramel-schmid@lemo.com"
        $Editor = "bpinot@lemo.com"
                
        #$StartDate = $files.CreationTime.Date.ToString("MM.dd.yyyy")

        $StartDate = $line.'Modification Date'.ToString("MM.dd.yyyy")
        $ModDate = $line.LastModified.ToString("MM.dd.yyyy")
        
        # Archived or not -----------------------------------
        [bool]$Archived = $false
		
        $NewFileName = $line.'FileName '.TrimEnd()+".pdf"
        # Replace "<BR>" by Carrege Return Line Feed
        #$desc1 = $line.'SP Descprition' -replace("<BR>","`r`n")

        $SmallName = $line.'FileName ' -replace($files.Extension,"")
        $Desc = @($line.F5, $line.F6, $line.F7, $line.F8) | Where-Object { $_ -ne "" }
        $Desc = if ($Desc -eq $null) { @("") } else { $Desc }


        Add-PnPFile -Path $Reference -Folder $DocLib -NewFileName $line.'FileName '.TrimEnd() -Values @{Author=$Author;Editor=$Editor;Created=$StartDate;Modified=$ModDate;
                     Title=$SmallName + " - " + $line.F5;
                     _ExtendedDescription=$Desc -join "<br/>"
                     Category=$line.F5
                     Customer=$line.F6 
                     
                     
                     } #> $null
    }

}