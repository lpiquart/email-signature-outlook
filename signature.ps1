#variables 
$SignatureName = 'Name signature' 
$SignatureVersion = "1.1"  
$ForceSignature = '0' #'0' = editable ; '1' non-editable and forced. 
  
#Environment variables 
$AppData=$env:appdata 
$SigPath = '\Microsoft\Signatures' 
$LocalSignaturePathold = $AppData+$SigPath + '.old'
$LocalSignaturePath = $AppData+$SigPath 
$RemoteSignaturePathFull = $SigSource 

#Get Active Directory information for logged in user 
$UserName = $env:username 
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 

#Test existance fichier signature
if (!(Test-Path -path $LocalSignaturePath'\'$SignatureName".htm" ))
{
    #Renomme Signature si ancienne version
    If (Test-Path -Path $LocalSignaturePath)
    {
        Rename-Item -path $LocalSignaturePath -newname $LocalSignaturePathold
        Write-Host "Dossier Signature renommé en signatures.old..." -ForegroundColor Yellow 

    }
    Write-Host "Création dossier signature" -ForegroundColor Yellow 
    New-Item $LocalSignaturePath -Type Directory | Out-Null
    New-Item -Path $LocalSignaturePath\$SignatureVersion -Type Directory| Out-Null
}
else
{
    #Test versionning
    If (!(Test-Path -Path $LocalSignaturePath\$SignatureVersion)) 
    {
        New-Item -Path $LocalSignaturePath\$SignatureVersion -Type Directory | Out-Null
    }
    else
    {
        Write-Host "Signature existe version $SignatureVersion, break ...." -ForegroundColor Yellow 
        break
    }
}

#CHOIX SIGNATURE SUR CHAMP IPPHONE
switch ($([string]($ADUser.ipphone)))
{
    "MODELE1"
        {
            $SigSource = '\\vm2012files\bureautique\1.Public\Procedures-Informatique\MODELE1.docx' 
            write-host "IPPHONE : MODELE1" -ForegroundColor Yellow 
        }
    "MODELE2"
        {
            $SigSource = '\\vm2012files\bureautique\1.Public\Procedures-Informatique\MODELE2.docx' 
            write-host "IPPHONE : MODELE2" -ForegroundColor Yellow 
        }
        default
        {
            write-host "IPPHONE non renseigné, break ...."
            break
        }
}

 
#Copy modèle signature
Write-Host "Copying Signatures" -ForegroundColor Green 
Copy-Item "$Sigsource" "$LocalSignaturePath\$SignatureName.docx" -Force 

#Insert variables from Active Directory to rtf signature-file 
$Word = New-Object -ComObject word.application 
$fullPath = "$LocalSignaturePath\$SignatureName.docx" 
$MSWord = $word.documents.open( $fullPath) 

#Remplacement lien hypertext
$hyperlinks = @($MSword.Hyperlinks) 
$hyperlinks | ForEach {
    If ($_.address -eq "https://cogeparc.eu/") {
    $_.address = "https://cogeparc.eu/index.php?email=" + "$([string]($ADUser.mail))"
    Write-Host "Attribut : https://cogeparc.eu/ changé en https://cogeparc.eu/index.php?email=$([string]($ADUser.mail))" -ForegroundColor yellow  
    }
}

#Remplace champ variable
$Word.Selection.Find.Execute("DisplayName", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.displayName))", 2) | Out-Null
Write-Host "Attribut : DisplayName changé par $([string]($ADUser.displayName))" -ForegroundColor yellow  
$Word.Selection.Find.Execute("department", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.department))", 2) | Out-Null
Write-Host "Attribut : Department changé par $([string]($ADUser.department))" -ForegroundColor yellow  
$Word.Selection.Find.Execute("title", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.title))", 2) | Out-Null
Write-Host "Attribut : title changé par $([string]($ADUser.title))" -ForegroundColor yellow 
$Word.Selection.Find.Execute("lienht", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.wWWhomepage))", 2) | Out-Null
Write-Host "Attribut : lienht changé par $([string]($ADUser.wWWhomepage))" -ForegroundColor yellow 

#REMPLACEMENT DES CHAMPS COMPLEMENTAIRES
switch ($([string]($ADUser.ipphone)))
{
    "MODELE1"
        {
            $Word.Selection.Find.Execute("TelephoneNumber", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.physicalDeliveryOfficeName))", 2) | Out-Null
            Write-Host "Attribut : TelephoneNumber changé par $([string]($ADUser.physicalDeliveryOfficeName))" -ForegroundColor yellow 
        }
    "MODELE2"
        {
            $Word.Selection.Find.Execute("TelephoneNumber1", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.telephonenumber))", 2) | Out-Null
            Write-Host "Attribut : TelephoneNumber1 changé par $([string]($ADUser.telephonenumber))" -ForegroundColor yellow 
            $Word.Selection.Find.Execute("TelephoneNumber2", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.homephone))", 2) | Out-Null
            Write-Host "Attribut : TelephoneNumber2 changé par $([string]($ADUser.telephonenumber))" -ForegroundColor yellow 
            $Word.Selection.Find.Execute("TelephoneNumber3", $false, $true, $false, $false, $false, $true, $FindContinue, $false, "$([string]($ADUser.mobile))", 2) | Out-Null
            Write-Host "Attribut : TelephoneNumber3 changé par $([string]($ADUser.mobile))" -ForegroundColor yellow 
        }
}


#Save new message signature  
 
#Save HTML 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML"); 
$path = $LocalSignaturePath+'\'+$SignatureName+".htm" 
$Word.ActiveDocument.saveas([ref]$path, [ref]$saveFormat) 
$Word.ActiveDocument.Close() 
$Word.Quit() 

#Suppression fichier.docx
remove-item "$LocalSignaturePath\$SignatureName.docx" 

If (Test-Path HKCU:'\Software\Microsoft\Office\16.0') 
{ 
    If ($ForceSignature -eq '0') 
    { 
    Write-host "Signature $SignatureName affectée en paramétrage outlook" -ForegroundColor Green 
 
    $Word = New-Object -comobject word.application 
    $EmailOptions = $Word.EmailOptions 
    $EmailSignature = $EmailOptions.EmailSignature 
    $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
    $EmailSignature.NewMessageSignature="$SignatureName" 
    $EmailSignature.ReplyMessageSignature="$SignatureName" 
 
    } 
    If ($ForceSignature -eq '1') 
    { 
        Write-Host "Ajout paramétrage Outlook forcé HKCU" 
        If (!(Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue))   
        {  
        New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force  
        }  
 
        If (!(Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue))   
        {  
        New-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force 
        }  
    } 
    else
    {
        Write-Host "Suppression paramétrage Outlook forcé HKCU" 
        If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)   
        {  
        Remove-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature'
        }  
 
        If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue)   
        {  
        Remove-ItemProperty HKCU:'\Software\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature'
        } 
    }
}
