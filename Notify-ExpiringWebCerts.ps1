[int]$TimeZoneOffset  = -5 # EST is -5
[int]$DaysToWarn      = 60
[string]$CA           = ""
[string]$SMTPserver   = ""
[string]$EmailTo      = ""
[string]$EmailFrom    = ""
[string]$EmailSubject = ""

Import-Module pspki
$issuedcerts = Get-CertificationAuthority $CA | Get-IssuedRequest -property *
$prettycerts = @()

foreach ($issuedcert in $issuedcerts) {

    if ($issuedcert.CommonName -ne $null -and $issuedcert.CertificateTemplate -eq 'WebServer' -and $issuedcert.NotAfter -gt (get-date) -and $issuedcert.NotAfter -lt (get-date).AddDays(60)) {
        $object = New-Object -TypeName PSObject
        $object | Add-Member –MemberType NoteProperty –Name RequestID –Value $issuedcert.RequestID
        $object | Add-Member –MemberType NoteProperty –Name CommonName –Value $issuedcert.CommonName
        $object | Add-Member –MemberType NoteProperty –Name Expires –Value $issuedcert.NotAfter.AddHours($TimeZoneOffset)
        $object | Add-Member –MemberType NoteProperty –Name CertificateTemplate –Value $issuedcert.CertificateTemplate
        $prettycerts = [array]$prettycerts + $object
    }
}

$EmailBody = "The following certificates are going to expire in the next $DaysToWarn days."
$EmailBody += $prettycerts | Format-Table -Wrap -AutoSize | Out-String

Send-MailMessage -SmtpServer $SMTPserver -To $EmailTo -From $EmailFrom -Subject $EmailSubject -Body $EmailBody
