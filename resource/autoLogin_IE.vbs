' ��FNAME,URL,ID,ID_TYPE,PW,PW_TYPE,SUBMIT(SUBMIT�ȊO�͕K�{���ڂł�)
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


' Shell�֘A�̑����񋟂���I�u�W�F�N�g���擾
Set sh = CreateObject("Shell.Application")
window_count = sh.Windows.Count
' MsgBox window_count

' IE�֘A�̑����񋟂���I�u�W�F�N�g���擾
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True

' ���������������������@�O�����@��������������������
arg_count = Wscript.Arguments.Count
'Wscript.Echo "count: " & arg_count

arg = Wscript.Arguments(0)
' Wscript.Echo "arg: " & arg
array = Split(arg, ",")
' Wscript.Echo "�v�f����" & UBound(array) +1 & "��"


' ���������������������@�A�N�Z�X�@��������������������
' URL�����
If array(1) <> "" Then
	oIE.Navigate(array(1))
	Wait()
Else
	Wscript.Echo "Error[1]: URL���i�[����Ă��܂���"
End If


' ���������������������@���O�C���@��������������������
If array(2) <> "" And array(3) <> "" Then
	' ���[�U�������
	oIE.document.getElementsByName(array(3)).Item(0).value = array(2)
Else
	Wscript.Echo "Error[2]: ID�܂���ID_TYPE���i�[����Ă��܂���"
End If

If array(4) <> "" And array(5) <> "" Then
	' �p�X���[�h�����
	oIE.document.getElementsByName(array(5)).Item(0).value = array(4)
Else
	Wscript.Echo "Error[3]: PW�܂���PW_TYPE���i�[����Ă��܂���"
End If

If array(6) <> "" Then
	' ���O�C���{�^�������
	oIE.document.getElementsByName(array(6)).Item(0).click
Else
    Wscript.Echo "Error[4]: SUBMIT���i�[����Ă��܂���"
End If

Select Case Err.Number
    Case 0
        WScript.Echo "����I�����܂����B"
    
    Case Else
        WScript.Echo "�\�����Ȃ��G���[���������܂����B" & vbCrLf & _
            "�G���[�ԍ��F" & Err.Number & vbCrLf & _
            "�G���[�ڍׁF" & Err.Description
End Select

Wait()


' ���������������������@�T�u���[�`���@��������������������
Function Wait()
	Do While oIE.Busy Or oIE.Readystate <> 4
		WScript.Sleep 100
	Loop
End Function