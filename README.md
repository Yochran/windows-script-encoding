# windows-script-encoding
A free and open source windows script host encoder for VBScript and JScript files.<br />
**I struggled to find nearly any documentation on the windows script encoder, so I created this VBScript which does it for you.**

# About
Encoding your windows script host scripts is an important part of deveoping windows script-based malware, otherwise the target can read exactly what it is you wrote, leaving you with the job of obfuscation.<br />
This script encoder will encode your VBScript and JScript files so that they are unreadable.

# Download
To download this encoder, click [here](https://github.com/Yochran/windows-script-encoding/releases/tag/1.0.0).<br />

# Usage
Once downloaded, run the encoder and fill out the options.<br />
**Options:**
  - Type (VBScript/JScript)
  - File (the input file, including it's file extension)
  - Encoded File (the name for the output encoded file, including it's encoded file extension, which is `.vbe` for VBScript, or `.jse` for JScript)


`This encoder must be in the same directory as the script you are encoding.`
