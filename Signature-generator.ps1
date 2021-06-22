#Email Signature Generator
#Author - Joshua Munro
#11/2020

$clientId = ""
$tenantName = ""
$clientSecret = ""
#contact Graph API 
$ReqTokenBody = @{
  Grant_Type    = "client_credentials"
  Scope         = "https://graph.microsoft.com/.default"
  client_Id     = "$clientid"
  Client_Secret = "$clientSecret"
}
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
$upn = whoami /upn
$apiUrl = 'https://graph.microsoft.com/v1.0/users/{0}' -f $upn
$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $apiUrl -Method Get

$displayname = $data.displayName
$jobtitle = $data.jobTitle
$businessphone = $data.mobilePhone
$Deskphone = $data.BusinessPhones
$mail = $data.mail
$officelocation = $data.officeLocation

$signatureFileName = "Signature.htm"
$ReplyFileName = "Reply.htm"
if ((Test-Path $env:APPDATA\microsoft\signatures) -eq $false)
    {
      new-item -ItemType directory -Name "Signatures" -Path $env:APPDATA\Microsoft\ -ErrorAction SilentlyContinue
    }

if($mail -like "*@*")
  {
    try{
    Copy-Item -Path "$PSScriptRoot\files" -Destination "$env:APPDATA\Microsoft\Signatures" -Force -Recurse
    Copy-Item -Path "$PSScriptRoot\Reply_files" -Destination "$env:APPDATA\Microsoft\Signatures" -Force -Recurse
  }
  catch{
    Write-Error "Unable to copy supporting files"
    exit
  }
}
else{
  write-error "This account contains no email"
}
#Main Signature
$html_template = @'
'@
$html_template = $html_template -replace 'DISPLAYNAMECHANGE', $displayname
$html_template = $html_template -replace 'EMAILCHANGE', $mail
$html_template = $html_template -replace 'JOBTITLECHANGE', $jobtitle
$html_template = $html_template -replace 'ADDRESSCHANGE', $officelocation
#$html_template = $html_template -replace 'PHONECHANGE',$businessphone
$html_template=if($businessphone -like "*[0-9]*")
  {
    $html_template -replace 'PHONECHANGE', "<p class=MsoNormal style='margin-right:-318.9pt;mso-pagination:none;mso-layout-grid-align:
    none;text-autospace:none'><span style='mso-ascii-font-family:Calibri;
    mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:#525252;
    mso-no-proof:yes'>Mobile: $businessphone<o:p></o:p></span></p>"
  } 
  Else
  {
  $html_template -replace 'PHONECHANGE', " "
  }
$html_template=if($deskphone -like "*[0-9]*")
  {
    $html_template -replace 'DESKCHANGE', "<p class=MsoNormal style='margin-right:-318.9pt;mso-pagination:none;mso-layout-grid-align:
    none;text-autospace:none'><span style='mso-ascii-font-family:Calibri;
    mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:#525252;
    mso-no-proof:yes'>Direct: $deskphone<o:p></o:p></span></p>"
  } 
  Else
  {
    $html_template -replace 'DESKCHANGE', " "
  }
$html_template | out-file $env:APPDATA\microsoft\signatures\$signatureFileName -Force

#reply template
$Reply_template = @'
'@

$Reply_template = $Reply_template -replace 'DISPLAYNAMECHANGE', $displayname
$Reply_template = $Reply_template -replace 'EMAILCHANGE', $mail
$Reply_template = $Reply_template -replace 'JOBTITLECHANGE', $jobtitle
$Reply_template = $Reply_template -replace 'ADDRESSCHANGE', $officelocation
#$Reply_template = $Reply_template -replace 'PHONECHANGE', $businessphone
$reply_template=if($businessphone -like "*[0-9]*")
  {
    $Reply_template -replace 'PHONECHANGE', "<p class=MsoNormal style='margin-right:-318.9pt;mso-pagination:none;mso-layout-grid-align:
    none;text-autospace:none'><span style='mso-ascii-font-family:Calibri;
    mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:#525252;
    mso-no-proof:yes'>Mobile: $businessphone<o:p></o:p></span></p>"
  }
  Else
  {
    $Reply_template -replace 'PHONECHANGE', " "
  }
$reply_template=if($deskphone -like "*[0-9]*")
  {
    $Reply_template -replace 'DESKCHANGE', "<p class=MsoNormal style='margin-right:-318.9pt;mso-pagination:none;mso-layout-grid-align:
    none;text-autospace:none'><span style='mso-ascii-font-family:Calibri;
    mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:#525252;
    mso-no-proof:yes'>Direct: $deskphone<o:p></o:p></span></p>"
  }
  Else
  {
    $Reply_template -replace 'DESKCHANGE', " "
  }
$Reply_template | out-file $env:APPDATA\microsoft\signatures\$ReplyFileName -Force
