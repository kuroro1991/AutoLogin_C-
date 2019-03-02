' 例：NAME,URL,ID,ID_TYPE,PW,PW_TYPE,SUBMIT(SUBMIT以外は必須項目です)
' arg(0):NAME/ arg(1):URL/ arg(2):ID/ arg(3):ID_TYPE/ arg(4):PW/ arg(5):PW_TYPE/ arg(6):SUBMIT

'Option Explicit
'On Error Resume Next

Dim sh
Dim window_count
Dim oIE
Dim new_oIE
Dim arg_count
Dim arg
Dim array


' Shell関連の操作を提供するオブジェクトを取得
Set sh = CreateObject("Shell.Application")
window_count = sh.Windows.Count
' MsgBox window_count

' IE関連の操作を提供するオブジェクトを取得
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True

' ＝＝＝＝＝＝＝＝＝＝　前処理　＝＝＝＝＝＝＝＝＝＝
arg_count = Wscript.Arguments.Count
'Wscript.Echo "count: " & arg_count

arg = Wscript.Arguments(0)
' Wscript.Echo "arg: " & arg
array = Split(arg, ",")
' Wscript.Echo "要素数は" & UBound(array) +1 & "個"


' ＝＝＝＝＝＝＝＝＝＝　アクセス　＝＝＝＝＝＝＝＝＝＝
' URLを入力
If array(1) <> "" Then
	oIE.Navigate(array(1))
	Wait()
Else
	Wscript.Echo "Error[1]: URLが格納されていません"
End If


' ＝＝＝＝＝＝＝＝＝＝　ログイン　＝＝＝＝＝＝＝＝＝＝
If array(2) <> "" And array(3) <> "" Then
	' ユーザ名を入力
	oIE.document.getElementsByName(array(3)).Item(0).value = array(2)
Else
	Wscript.Echo "Error[2]: IDまたはID_TYPEが格納されていません"
End If

If array(4) <> "" And array(5) <> "" Then
	' パスワードを入力
	oIE.document.getElementsByName(array(5)).Item(0).value = array(4)
Else
	Wscript.Echo "Error[3]: PWまたはPW_TYPEが格納されていません"
End If

If array(6) <> "" Then
	' ログインボタンを入力
	oIE.document.getElementsByName(array(6)).Item(0).click
Else
    Wscript.Echo "Error[4]: SUBMITが格納されていません"
End If

Select Case Err.Number
    Case 0
        WScript.Echo "正常終了しました。"
    
    Case Else
        WScript.Echo "予期しないエラーが発生しました。" & vbCrLf & _
            "エラー番号：" & Err.Number & vbCrLf & _
            "エラー詳細：" & Err.Description
End Select

Wait()


' ＝＝＝＝＝＝＝＝＝＝　サブルーチン　＝＝＝＝＝＝＝＝＝＝
Function Wait()
	Do While oIE.Busy Or oIE.Readystate <> 4
		WScript.Sleep 100
	Loop
End Function