###

#region Initialisation
$Dir = "C:\Users\Administrator\Documents\Checkpoint2Files"
$CsvFile = "Users.csv"
$NecessaryFiles = ($CsvFile,"fonction_Random-Password.txt")
$UsersOUName = "Utilisateurs"

$Societe = Read-Host "Donnez le nom de votre Societe"
$Domain = Read-Host "Donnez le nom de votre Domaine"
####################mdp
Function Random-Password ($length = 23)
{

    $punc = 42..42
    $digits = 48..57
    $letters = 65..90 + 97..122

    $password = get-random -count $length -input ($punc + $digits + $letters) |`
    ForEach -begin { $aa = $null } -process {$aa += [char]$_} -end {$aa}
   

    Return $password.ToString()
}




#############################
Import-Module ActiveDirectory
#endregion

#region Cr ation dossiers et copie de fichiers
If (-not(Test-Path $Dir))
{
    New-Item -Path $Dir -ItemType Directory -Force:$false
}
foreach ($File in $NecessaryFiles){Copy-Item -Path ".\$File" -Destination $Dir -Recurse -Force}
#endregion

### import du .csv
$GcFileCsv = Import-Csv -Path "$Dir\$CsvFile" -Delimiter ";" -Header "prenom","nom","societe","fonction","service","description","mail","mobile","scriptPath","telephoneNumber" -Encoding UTF7

#region OU
### si l'ou n'existe pas
If (-not(Get-ADOrganizationalUnit -Filter {Name -eq $UsersOUName}))
{
    New-ADOrganizationalUnit -Path "dc=$societe,dc=$Domain" -Name "$UsersOUName" -Confirm:$False -ProtectedFromAccidentalDeletion:$false
}
### List d'OU cr e   partir de la colonne service dans le .csv dans l'OU utilisateur
$OUList = ($GcFileCsv | Select service -Unique ).service 
Foreach ($Ou in $OuList)
{
    If(-not(Get-ADOrganizationalUnit -Filter {Name -eq $Ou} -SearchBase "OU=$UsersOUName,DC=$societe,DC=$Domain"))
    {
        if ($Ou -eq "service"){
        }

        else {

                 Write-Host "$Ou existe pas"
                 New-ADOrganizationalUnit -Path "ou=$UsersOUName,dc=$societe,dc=$Domain" -Name "$Ou" -Confirm:$False -ProtectedFromAccidentalDeletion:$false
                 Write-Host "$Ou ==> Cr ation"}
    }

}
#endregion
                  
### met dans une variable le nom samAccountName d'un utilisateur
$ADUsers = Get-ADUser -Filter * -Properties SamAccountName
### boucle qui parcours les utilisateur dans le csv
Foreach ($User in $GcFileCsv)
{
    $SamAccountName = "$($User.prenom).$($User.nom)"
    $ResCheckADUser = $ADUsers | Where {$_.SamAccountName -eq $SamAccountName}
    If ($ResCheckADUser -eq $null)
    {
        ### On va cr e le user avec les info suivant
        Write-Host "Compte $SamAccountName   Cr er"

        $Password          = Random-Password
        $Password          = $Password + "*"
        
        
        $MobilePhone       = $User.mobile
        $OfficePhone       = $User.telephoneNumber
        $Title             = $User.fonction
        $Department        = $User.service
        $Description       = "Utilisateur du service $Department - $title"
        $ScriptPath        = $User.scriptPath
        $UserPrincipalName = "$SamAccountName@$societe.$Domain"
        $Name              = "$($User.prenom) $($User.nom)"
        $GivenName         = $User.prenom
        $Surname           = $User.nom
        $DisplayName       = $SamAccountName
        $Path              = "OU=$($User.service),OU=$UsersOUName,DC=$societe,DC=$Domain"
        $Company           = $User.societe
        $Email             = $User.mail
        


        ### cr e l'user avec les infos contenu dans les variable
        New-ADUser `
            -SamAccountName $SamAccountName `
            -UserPrincipalName $UserPrincipalName `
            -Name $Name `
            -GivenName $GivenName `
            -Surname $Surname `
            -Enabled $True `
            -DisplayName $DisplayName `
            -Path $Path `
            -Company $Company `
            -OfficePhone $OfficePhone `
            -MobilePhone $MobilePhone `
            -EmailAddress $Email `
            -Title $Title `
            -Department $Department `
            -AccountPassword (ConvertTo-SecureString "$Password" -AsPlainText -force) `
            -ChangePasswordAtLogon $True `
            -Description $Description `
            -ScriptPath $ScriptPath
        Write-Host "$Name ==> cr e avec le mot de passe $password" -ForegroundColor Green
    }
    Else
    {
        ### dans le cas ou l'user existe deja
        Write-Host "$Name d ja cr e" -ForegroundColor Red -BackgroundColor Yellow
    }
}