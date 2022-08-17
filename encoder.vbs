Option Explicit

' --
' Encode function
' @params:
' - sInput: The inputted file
' - sExtension: The file extension
' - sType: The file type
' - sOutput: The output encoded file
' --
Function Encode(sInput, sExtension, sType, sOutput)
    ' Get encoder and FSO
    dim encoder, fso
    set encoder = CreateObject("Scripting.Encoder")
    set fso = CreateObject("Scripting.FileSystemObject")

    ' Get input file as stream
    dim input_file, file_stream, source_file
    set input_file = fso.GetFile(sInput)
    set file_stream = input_file.OpenAsTextStream(1)
    source_file = file_stream.ReadAll
    file_stream.Close

    ' Encode file
    dim encoded_file
    set encoded_file = fso.CreateTextFile(sOutput)
    encoded_file.Write encoder.EncodeScriptFile(sExtension, source_file, 0, sType)
    encoded_file.Close
End Function

' Get arguments
Dim arg1, arg2, arg3, arg4, box
arg1 = WScript.Arguments(0)
arg2 = WScript.Arguments(1)
arg3 = WScript.Arguments(2)
arg4 = WScript.Arguments(3)

' Run encode function
Encode arg1, arg2, arg3, arg4