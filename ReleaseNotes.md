# OutSpeech Release Notes

OutSpeech is a PowerShell module inspired by inspired by these articles:

- [Give PowerShell a Voice](https://learn-powershell.net/2013/12/04/give-powershell-a-voice-using-the-speechsynthesizer-class/)
- [OutVoice](https://gallery.technet.microsoft.com/scriptcenter/Out-Voice-1be16d5e)

OutSpeech provides the following features:

- An Out-Speech function for sending PowerShell output to the default audio device
- An Export-Speech function for sending PowerShell output to wav file
- Functions to Enable and Set Speech Configuration objects to use/re-use with Out-Speech, enabling the use of different configurations for output Rate, Volume, and Voice
- Get-SpeechVoice function to enumerate the available voices installed on a system
- a Disable-SpeechConfiguration function for cleanup of the Speech Configuration Object(s) used or created by OutSpeech.

## OutSpeech 1.0.0.0

First published version in the PSGallery
