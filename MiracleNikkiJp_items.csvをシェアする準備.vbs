Dim src_filename
Dim filesys
Dim cwd_name
Dim src_stream
Dim dst_stream
Dim ch

' Determine the path name of the file to be accessed
src_filename = "MiracleNikkiJp_items.csv"
set filesys = CreateObject("Scripting.FileSystemObject")
if InStr(src_filename, "\") = 0 then
    ' current directory
    cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
    src_filename = cwd_name & "\" & src_filename
end if

' Open source stream
set src_stream = CreateObject("ADODB.Stream")
' Check charset
src_stream.Type = 1  ' binary
src_stream.Open
src_stream.LoadFromFile src_filename
ch = src_stream.Read(1)
if Ascb(ch) = &hEF then
    ' already utf-8
    src_stream.Close
    ' message
    WScript.echo("MiracleNikkiJp_items.csvのシェアの準備は既にできているようです。何もせずに終了します。")
    WScript.Quit(-1)
end if
' message
WScript.echo("MiracleNikkiJp_items.csvの文字列をutf-8に、タブをカンマに変換します。")
' rewind and text mode
src_stream.Position = 0
src_stream.Type = 2  ' text
src_stream.Charset = "shift_jis"
' Open destination stream
set dst_stream = CreateObject("ADODB.Stream")
dst_stream.Type = 2  ' text
dst_stream.Charset = "utf-8"
dst_stream.Open
' read whole
src_str = src_stream.ReadText(-1)
' replace tab to ,
src_str = Replace(src_str, "	", ",")
' write to stream
dst_stream.WriteText src_str, 0
' Close source stream
src_stream.Close
' Overwrite with utf8 code with BOM
dst_stream.SaveToFile src_filename, 2
' Close destination stream
dst_stream.Close
' message
WScript.echo("MiracleNikkiJp_items.csvのシェアの準備が終了しました。")
