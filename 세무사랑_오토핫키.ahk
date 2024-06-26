#Requires AutoHotkey v2.0

main := Gui(Options := 'Resize -SysMenu',"세무사랑 오토핫키")
radio1 :=main.AddRadio('Center vTrigger',"법인")
radio2 :=main.AddRadio('Center',"개인")
radio1.OnEvent('Click',법인활성화)
radio2.OnEvent('Click',개인활성화)
main.Show("AutoSize")

;쓰려는 엑셀 켜놔야 함
xl := ComObjActive("excel.application")

;중지 키 
#SuspendExempt true
^w:: Suspend
#SuspendExempt false

;법인개인 버튼 작동
법인활성화(*){
    if (radio1 := 1){
        global 설정 := "법인"
    }   
}

개인활성화(*){
    if (radio2 := 1){
        global 설정 := "개인"
    }   
}



::1:: ;매크로 설정할 키
{
    지정 := xl.Sheets("설정").Range("A:A").find("1") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::2:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("2") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::3:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("3") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::4:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("4") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::5:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("5") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::6:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("6") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::7:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("7") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::8:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("8") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::9:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("9") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}

::10:: ;매크로 설정할 키
{
    지정 := xl.Sheets(설정).Range("A:A").find("10") ;find 안에 값 수정하면 원하는 값 찾음
    번호 := 지정.offset(0,1).Text ;지정에서 찾은거 옆으로 한칸 있는 값
    SendInput(번호)
    Send('{Enter}')
}
