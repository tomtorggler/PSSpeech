
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
        [Parameter()]
        [ValidateSet("westeurope","northeurope","westus2","westus","uksouth","francecentral")]
        [string]
        $Region = "westeurope",

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Key 
    )
    $FetchTokenHeader = @{
        'Content-type'='application/x-www-form-urlencoded';
        'Content-Length'= '0';
        'Ocp-Apim-Subscription-Key' = $Key
    } 
    New-Object -TypeName psobject -Property (@{
        TimeStamp = Get-DAte
        Token = Invoke-RestMethod -Method POST -Uri "https://$region.api.cognitive.microsoft.com/sts/v1.0/issueToken" -Headers $FetchTokenHeader
    })
}

function Save-SpeechToken {
    <#
    .SYNOPSIS
        Save a token for the current session.
    .DESCRIPTION
        This function takes a token as retreived from Get-SpeechToken and creates a variable in the global scope and saves the token.
    .EXAMPLE
        PS C:\> Get-SpeechToken -Key <yourkey> | Save-SpeechToken

        This example first gets a token then saves it to a global variable in the current PowerShell session.
    .EXAMPLE
        PS C:\> Get-SpeechToken -Key <yourkey> -OutVariable token
        PS C:\> Save-SpeechToken -Token $token

        This example first gets a token then saves it to a global variable in the current PowerShell session.
    .INPUTS
        [psobject]
    .OUTPUTS
        None.
    #>
    [CmdletBinding(HelpUri = 'https://ntsystems.it/PowerShell/Save-SpeechToken/')]
    param (
        [Parameter(ValueFromPipeline)]
        $Token    
    )
    Set-Variable -Scope global -Name PSSpeechToken -Value $token
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
        [ValidateSet("westeurope","northeurope","westus2","westus","uksouth","francecentral")]
        [string]
        $Region = "westeurope",

        [ValidateNotNullOrEmpty()]
        $Token = $Global:PSSpeechToken
    )
    $AuthHeader = @{
        'Content-type' = 'application/ssml+xml';
        'Authorization' = "Bearer $($Token.Token)";
        'Content-Length'= '0';
    }   
    Invoke-RestMethod -uri "https://$region.tts.speech.microsoft.com/cognitiveservices/voices/list" -Headers $AuthHeader -Method Get
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
        [Parameter()]
        [ValidateSet("westeurope","northeurope","westus2","westus","uksouth","francecentral")]
        [string]
        $Region = "westeurope",

        [Parameter()]
        $Token = $Global:PSSpeechToken,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Text,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        [Parameter()]
        [ValidateSet('en-US-GuyNeural','en-US-JessaNeural','zh-CN-XiaoxiaoNeural','it-IT-ElsaNeural','de-DE-KatjaNeural','en-GB-HarryNeural','fr-FR-HortenseNeural','pt-BR-FranciscaNeural')]
        [string]
        $Voice = 'en-GB-HarryNeural',
        
        [Parameter()]
        [ValidateSet("raw-16khz-16bit-mono-pcm","audio-16khz-128kbitrate-mono-mp3","audio-16khz-32kbitrate-mono-mp3","audio-24khz-96kbitrate-mono-mp3","audio-24khz-48kbitrate-mono-mp3","audio-24khz-160kbitrate-mono-mp3","audio-16khz-64kbitrate-mono-mp3")]
        [string]
        $OutputFormat = "audio-16khz-32kbitrate-mono-mp3"
    )
    $AuthHeader = @{
        'Content-type' = 'application/ssml+xml'
        'Authorization' = "Bearer $($Token.Token)"
        'X-Microsoft-OutputFormat' = $OutputFormat
        'User-Agent' = "powershell"
    }
    # build the ssml xml 
    [xml]$xml = "<speak version='1.0' xml:lang='en-GB'><voice xml:lang='en-GB' xml:gender='Female' name='$Voice'>$Text</voice></speak>"
    # send to speech service and save output in file 
    Invoke-RestMethod -Uri "https://$region.tts.speech.microsoft.com/cognitiveservices/v1" -Headers $AuthHeader -Method Post -Body $xml -OutFile $Path
}

