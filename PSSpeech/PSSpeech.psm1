function Get-SpeechToken {
    <#
    .SYNOPSIS
        Get OAuth token for authorization to Azure Cognitive Services.
    .DESCRIPTION
        This function uses Invoke-RestMethod to get a bearer token that can be used in the Authorization header when calling 
        Azure Cognitive Services. This requires access to an Azure subscription and API key for the speech service. 
    .EXAMPLE
        PS C:\> Get-SpeechToken -Key <yourkey> 

        This example gets a token using the provided key. The default value for the Region parameter is set to westeurope, please specify the region where your Cognitive Services is deployed.
    .INPUTS
        None.
    .OUTPUTS
        [psobject]
    .NOTES
        Key should probably be a secure string, update once secrets management module is released.   
    #>
    [CmdletBinding(HelpUri = 'https://ntsystems.it/PowerShell/Get-SpeechToken/')]
    param (
        [Parameter(mandatory=$true)]
        [ValidateSet("centralus", "eastus", "eastus2", "northcentralus", "southcentralus", "westcentralus", "westus", "westus2", "canadacentral", "brazilsouth", "eastasia", "southeastasia", "australiaeast", "centralindia", "japaneast", "japanwest", "koreacentral", "northeurope", "westeurope", "francecentral", "uksouth")]
        [string]
        $Region,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Key 
    )
    $FetchTokenHeader = @{
        'Content-type'='application/x-www-form-urlencoded';
        'Content-Length'= '0';
        'Ocp-Apim-Subscription-Key' = $Key
    } 
    $script:SpeechToken = New-Object -TypeName psobject -Property (@{
        TimeStamp = Get-Date
        Token = Invoke-RestMethod -Method POST -Uri "https://$region.api.cognitive.microsoft.com/sts/v1.0/issueToken" -Headers $FetchTokenHeader
        Region = $Region
    })
    
}

function Get-SpeechTokenResult
{
    return $script:SpeechToken
}

function Get-SpeechVoicesList {
    <#
    .SYNOPSIS
        Get a list of available voices from the speech service.
    .DESCRIPTION
        This function uses Invoke-RestMethod to get a list of available voices from the Azure Cognitive Services Speech Service. Use the Token parameter
        to specify a token created with Get-SpeechToken and use the Region parameter to specify a region other than the default westeurope.
        If the Token parameter is not specified, the global variable created by Save-SpeechToken is used.
    .EXAMPLE
        PS C:\> Get-SpeechVoicesList

        This example gets a list of available voices.
    .INPUTS
        None.
    .OUTPUTS
        [psobject]
    .NOTES
        General notes
    #>
    [CmdletBinding(HelpUri = 'https://ntsystems.it/PowerShell/Get-SpeechVoicesList/')]
    param (
    )
    $token = $script:SpeechToken.Token
    $AuthHeader = @{
        'Content-type' = 'application/ssml+xml';
        'Authorization' = "Bearer $Token";
        'Content-Length'= '0';
    }   
    $region = $script:SpeechToken.region
    $voices = Invoke-RestMethod -uri "https://$region.tts.speech.microsoft.com/cognitiveservices/voices/list" -Headers $AuthHeader -Method Get
    $voices
}

function Convert-TextToSpeech {
    <#
    .SYNOPSIS
        Convert a string to audio using Azure Cognitive Services. 
    .DESCRIPTION
        This function uses Invoke-RestMethod to call the Azure Cognitive Service Speech Service API, convert a string to speech, and save the resulting audio to a file.
    .EXAMPLE
        PS C:\> Convert-TextToSpeech -Text "Hi, this is a test." -Path test.mp3

        This example converts the string "Hi, this is a test." to speech and saves the audio to the test.mp3 file.
    .INPUTS
        None.
    .OUTPUTS
        None.
    .NOTES
        I've added only the neural voices to the ValidateSet attribute, more voices are available.
    #>
    [CmdletBinding(HelpUri = 'https://ntsystems.it/PowerShell/Convert-TextToSpeech/')]
    param (
        
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Text,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        [Parameter()]
        [ValidateScript({Get-SpeechVoicesList | select-object ShortName})]
        [string]
        $Voice = 'en-GB-LibbyNeural',
        
        [Parameter()]
        [ValidateSet("raw-16khz-16bit-mono-pcm","audio-16khz-128kbitrate-mono-mp3","audio-16khz-32kbitrate-mono-mp3","audio-24khz-96kbitrate-mono-mp3","audio-24khz-48kbitrate-mono-mp3","audio-24khz-160kbitrate-mono-mp3","audio-16khz-64kbitrate-mono-mp3")]
        [string]
        $OutputFormat = "audio-16khz-32kbitrate-mono-mp3"
    )
    $SpokenVoice = Get-SpeechVoicesList | ? {$_.ShortName -eq $Voice}
    $token =$script:SpeechToken.Token
    $AuthHeader = @{
        'Content-type' = 'application/ssml+xml'
        'Authorization' = "Bearer $token"
        'X-Microsoft-OutputFormat' = $OutputFormat
        'User-Agent' = "powershell"
    }
    $region = $script:SpeechToken.region
    # build the ssml xml 
    #<voice xml:lang='en-US' xml:gender='Female'     name='en-US-AriaRUS'>
    #will assume language and voice are the same
    [xml]$xml = "<speak version='1.0' xml:lang='en-GB'><voice xml:lang='en-GB' xml:gender='Female' name='$Voice'></voice></speak>"
    $xml.speak.voice.InnerText=$text
    $xml.speak.lang = $spokenvoice.Locale
    $xml.speak.voice.lang = $spokenvoice.Locale
    $xml.speak.voice.gender = $spokenvoice.Gender
    $xml.speak.voice.name = $spokenvoice.ShortName
    # send to speech service and save output in file 
    Invoke-RestMethod -Uri "https://$region.tts.speech.microsoft.com/cognitiveservices/v1" -Headers $AuthHeader -Method Post -Body $xml -OutFile $Path -Verbose
}

