Dim iFile As Integer
Dim sLine As String
Dim ar() As String

Dim x As Double
Dim y As Double
Dim z As Double
Dim s As String

iFile = FreeFile

Open "c:\myfile.txt" For Input As iFile

Do

    Line Input #iFile, sLine
   
    ar = Split(sLine, ",")
   
    If UBound(ar) = 3 Then
        x = Val(ar(0))
        y = Val(ar(1))
        z = Val(ar(2))
        s = ar(3)
        ' process point...
        Debug.Print x, y, z, s
    End If

Loop Until EOF(iFile)

Close iFile
