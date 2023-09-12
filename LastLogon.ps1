Import-Module ImportExcel
Import-Module ActiveDirectory

$utilisateurs = Get-ADUser -Filter * -Properties lastLogon

$resultats = @()

foreach ($utilisateur in $utilisateurs) {
    $nomUtilisateur = $utilisateur.SamAccountName
    $lastLogon = [DateTime]::FromFileTime($utilisateur.lastLogon)

    $ligneFormattee = [PSCustomObject]@{
        "Nom d'utilisateur" = $nomUtilisateur
        "LastLogon" = $lastLogon.ToString("yyyy-MM-dd HH:mm:ss")
    }

    $resultats += $ligneFormattee
}

$resultats | Export-Excel -Path "C:\Users\adm-jbo\Desktop\lastLogon_AD.xlsx" -AutoSize -AutoFilter

Write-Host "Les dates de lastLogon de tous les utilisateurs de l'Active Directory ont été exportées dans un fichier Excel"
