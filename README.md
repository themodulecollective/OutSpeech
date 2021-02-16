# OutSpeech

OutSpeech is a Windows PowerShell module inspired by inspired by these articles:

- [Give PowerShell a Voice](https://learn-powershell.net/2013/12/04/give-powershell-a-voice-using-the-speechsynthesizer-class/)
- [OutVoice](https://gallery.technet.microsoft.com/scriptcenter/Out-Voice-1be16d5e)

OutSpeech provides the following features:

- An Out-Speech function for sending PowerShell output to the default audio device
- An Export-Speech function for sending PowerShell output to wav file
- Functions to Enable and Set Speech Configuration objects to use/re-use with Out-Speech, enabling the use of different configurations for output Rate, Volume, and Voice
- Get-SpeechVoice function to enumerate the available voices installed on a system
- a Disable-SpeechConfiguration function for cleanup of the Speech Configuration Object(s) used or created by OutSpeech.

## Release Notes

1.0.3 Documentation Updates : Updates to this readme.md file.
1.0.2 License and metadata updates.
1.0.1 Original release to PSGallery


## How it Works

1. NOTE: Only works on Windows machines right now (no .Net core support for speech synthesis yet).  Import the OutSpeech module into your Windows PowerShell session (if using PowerShell 6.x or 7.x, use the Import-Module parameter -UseWindowsPowerShell).
2. Simple usage: 'Out-Speech "Hello, Dave"'
   - Out-Speech will create a Default SpeechConfiguration object
   - Out-Speech will re-use the Default SpeechConfiguration object if it exists
   - the DefaultSpeech object will have default voice, rate, and volume settings for your system (these can be overridden by the Rate, Volume, or Voice parameters of Out-Speech)
   - Out-Speech should speak the specified text
3. More Advanced Usage:
   - create a speech object profile with Enable-SpeechConfiguration, specifying a ConfigurationName, a rate, a voice, and a volume setting to use
   - use the SpeechConfiguration Out-Speech by specifying the ConfigurationName
   - modify a SpeechConfiguration using Set-SpeechConfiguration
   - create a wav file of speech using Export-Speech
4. A note about the rate, volume, and voice parameters when used with Out-Speech
    -When you use the Rate, Volume, or Voice parameters with Out-Speech the settings will persist on the SpeechConfiguration object specified (or the Default SpeechConfiguration)

## Development Plans

- Export-Speech: Support [Audio Format Info](https://msdn.microsoft.com/en-us/library/ms586885(v=vs.110).aspx)
- Add support for SSML with the [SpeakSsml](http://www.w3.org/TR/speech-synthesis/) and/or [SpeakSsmlAsync](https://msdn.microsoft.com/en-us/library/office/hh361578(v=office.14).aspx) methods.  These let you manipulate the voice/reading for better results.
- [Speech Recognition?](https://gallery.technet.microsoft.com/scriptcenter/Fun-with-PowerShell-and-c59c3d4b#content)

## Basic Usage Example

```PowerShell
    Import-Module OutSpeech
    Out-Speech "Hello, Dave"
    Out-Speech "This is Microsoft Zira Desktop" -Rate 3 -Volume 90 -Voice "Microsoft Zira Desktop"
```

This is one of the default voices on my Windows 10 system, yours may vary depending on language packs/editions of Windows. Use Get-SpeechVoice to see what is available on your local system.

## Advanced Usage Example

Configuring multiple SpeechConfigurations to use on demand.  Passthru parameter allows you to view the resulting object.

```PowerShell
    Import-Module OutSpeech -force
    Enable-SpeechConfiguration -ConfigurationName "DavidFastLoud" -Rate 3 -Voice "Microsoft David Desktop" -Volume 100
    Enable-SpeechConfiguration -ConfigurationName "ZiraSlowSoft" -Rate -1 -Voice "Microsoft Zira Desktop" -Volume 20
    Out-Speech -ConfigurationName ZiraSlowSoft -inputobject "I am speaking to you slowly and not very loudly" -Synchronous
    "I am speaking to you loudly and at a fast rate" | Out-Speech -ConfigurationName DavidFastLoud -Synchronous
    Set-SpeechConfiguration -ConfigurationName DavidFastLoud -Rate 10
    Out-Speech -inputobject "Now I am speaking to you at my maximum rate." -ConfigurationName DavidFastLoud
    Disable-SpeechConfiguration -ConfigurationName 'DavidFastLoud','ZiraSlowSoft'
```

## License

OutSpeech is released under the MIT License
