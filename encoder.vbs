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

' Get type
Dim sType, sExtension
sType = InputBox("Enter the type (VBScript, JScript):")

If StrComp(sType, "VBScript") Then
    sExtension = ".vbs"
Else if StrComp(sType, "JScript") Then
    sExtension = ".js"
    End If
End If

' Get file
Dim sInput
sInput = InputBox("Enter the name of the file you wish to encode:")

' Get output
Dim sOutput
sOutput = InputBox("Enter the name of the output file (including encoded extension):")

' Check if length = 0, if not, encode file
If Len(sInput) = 0 OR Len(sExtension) = 0 OR Len(sType) = 0 OR Len(sOutput) = 0 Then
    Dim sErr
    sErr = MsgBox("One or more inputs we're not entered.",0+16,"Error")
Else
    ' Encode script
    Encode sInput, sExtension, sType, sOutput
End If