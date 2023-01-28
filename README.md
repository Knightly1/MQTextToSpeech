# MQTextToSpeech

Uses Microsoft's Speech API to provide text to speech options for outputting speech.

## Getting Started

```txt
/plugin MQTextToSpeech
/tts say If you give a man a fire he's warm for a day.  But if you set a man on fire, he's warm for the rest of his life.
/mqsettings plugins/TextToSpeech
```

### Commands

Primarily the command `/tts say` is used to speak text.  Other commands available are below, but since they are mostly for settings
it is better to use the settings menu.

```txt
Usage:
    /tts loglevel|reload|say|sayxml|speed|voice|volume [argument]
Available Verbs:
    LogLevel - Set the log level (default is 'info')
    Reload - Reload the settings from the ini file
    Say - Speak a line of text
    SayXML - Speak a line of text using XML formatting for pattern changes
    Speed - Set the speech speed (-10 to 10)
    Voice - Change the voice (must be installed)
    Volume - Set the volume of speech (0 to 100)
  NOTE: All settings are available in the settings menu.
Example:
    /tts say Some people juggle geese.
```

### TLOs and Members

```
TLO and Members:
    ${TTS.Voice} -- Which voice is being used
    ${TTS.Volume} -- What volume is currently set
    ${TTS.Speed} -- What speech speed is currently set
Example:
    /tts say The current voice is ${TTS.Voice}.
```

### Configuration File

MQTextToSpeech.ini stores the persistent settings (Volume/Speed/Voice).  The voice is separate between 32 bit and 64 bit since they
maybe be different (even on the same system).

```ini
[Settings]
Volume=100
Speed Modifier=0
Voice=Microsoft David
Voice_32bit=Microsoft Zira
```

### Using XML
TextToSpeech supports the SAPI XML TTS API when using the `SayXML` parameter.  To find more about how to use this to adjust pitch, rate, and speed within a single line
you can review the api here: https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms717077(v=vs.85)

For example, the empty box in settings can be recreated using this command:
```
/say sayxml You <pitch middle="10">really should enter</pitch> something to say if you want to test.  This box also supports <spell>XML</spell>
```

## Authors

* **Knightly** - *Initial work*

## Acknowledgements

* **A_Enchanter_00** - Wrote MQ2VoiceCommand (which doesn't actually do voice commands, sadly)
* **eqholic** - Wrote MQ2Speaker which wasn't ported to MQ
