# OutSpeech

  OutSpeech is a PowerShell module inspired by inspired by https://learn-powershell.net/2013/12/04/give-powershell-a-voice-using-the-speechsynthesizer-class/ and https://gallery.technet.microsoft.com/scriptcenter/Out-Voice-1be16d5e
  OutSpeech provides the following features:

* An Out-Speech function for sending PowerShell output to the default audio device
* An Export-Speech function for sending PowerShell output to wav file
* Functions to Enable and Set Speech Configuration objects to use/re-use with Out-Speech, enabling the use of different configurations for output Rate, Volume, and Voice
* Get-InstalledVoice function to enumerate the available voices installed on a system
* a Disable-Speech function for cleanup of the Speech Configuration Object(s) used or created by OutSpeech.

## How it Works
1. import the OutSpeech module into your PowerShell session 
2. Simple usage: 'Out-Speech "Hello, Dave"'
    - Out-Speech will create a DefaultSpeech object
    - Out-Speech will re-use the DefaultSpeech object if it exists
    - the DefaultSpeech object will have default voice, rate, and volume settings for your system (these can be overridden by the Rate, Volume, or Voice parameters of Out-Speech)
    - Out-Speech should play the specified text
3. More Advanced Usage:
    - create a speech object profile with Enable-Speech, specifying a SpeechConfiguration name, a rate, a voice, and a volume setting to use
    - use the speech object configured by Enable-Speech with Out-Speech by specifying the SpeechConfiguration name you used with Enable-Speech
    - modify the speech object settings using Set-Speech
    - re-use the speech object with the modified settings by using Out-Speech again with the same SpeechConfiguration name
4. A note about the rate, volume, and voice parameters when used with Out-Speech
    -When you use the Rate, Volume, or Voice parameters with Out-Speech the settings will persist on the SpeechConfiguration object specified (or the DefaultSpeech object)
    -unless you also use the SynchronousOutput and RestoreSpeechObjectSettings parameters with Out-Speech
    -This is because when output is asynchronous we would have to set up event handling to wait/watch for output to complete before resetting the Rate, Volume, or Voice attributes of the object back to their original settings
    -In effect, this would make the function behave as if it were synchronous
    -So, it is a best practice to not set the Rate, Volume, or Voice attributes directly with Out-Speech in most cases.  Rather, you should do that when you run Enable-Speech, or you should do it with Set-Speech
    -On the other hand, for one-off changes to Rate, Volume, or Voice attributes when you can deal with synchronous audio rendering, you can do it with Out-Speech.
    -My goal was to keep it as flexible as possible but to make the user aware of the synchronous/asynchronous issue when using the SpeechConfiguration objects.  

## Development Plans
1. Make the -voice parameter of each function that has it a dynamic parameter with currently installed voices or add validation within the function for voice specified
2. Export-Speech: Support Audio Format Info:  https://msdn.microsoft.com/en-us/library/ms586885(v=vs.110).aspx
3. Add inline documentation to functions
4. Add support for SSML with the SpeakSsml and/or SpeakSsmlAsync methods: http://www.w3.org/TR/speech-synthesis/ and https://msdn.microsoft.com/en-us/library/office/hh361578(v=office.14).aspx

## Basic Usage Example
    Import-Module OutSpeech
    Out-Speech "Hello, Dave"
    Out-Speech "This is Microsoft Zira Desktop" -Rate 3 -Volume 90 -Voice "Microsoft Zira Desktop" 
This is one of the default voices on my Windows 10 system, yours may vary depending on language packs/editions of Windows.
If your WarningPreference is set to Continue, you would see a Warning with the last command above since you would be modifying the DefaulSpeech configuration object with changes that will persist on the object.
The following command would avoid the warning, but requires Synchronous output to your default audio device

    Out-Speech "This is Microsoft David Desktop" -voice "Microsoft David Desktop" -SynchronousOutput -RestoreSpeechObjectSettings

## Advanced Usage Example
Configuring multiple SpeechConfigurations to use on demand.  Passthru parameter allows you to view the resulting object.

    Import-Module OutSpeech
    Enable-Speech -SpeechConfiguration "DavidFastLoud" -Rate 6 -Voice "Microsoft David Desktop" -Volume 100 -passthru
    Enable-Speech -SpeechConfiguration "ZiraSlowSoft" -Rate -1 -Voice "Microsoft Zira Desktop" -Volume 20 -passthru
    Out-Speech -SpeechConfiguration ZiraSlowSoft "I am speaking to you slowly and not very loudly"
    "I am speaking to you loudly and at a very high rate" | Out-Speech -SpeechConfiguration DavidFastLoud
    Set-Speech -SpeechConfiguration DavidFastLoud -Rate 10
    Out-Speech -inputobject "Now I am speaking to you at my maximum rate." -SpeechConfiguration DavidFastLoud

## License

OutSpeech is released under the MIT License

