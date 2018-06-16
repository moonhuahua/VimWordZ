;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;作者：Space；QQ：105224583
;以下脚本为本人去掉个人敏感信息后余下的完整脚本，希望能对ahk爱好者提供一点学习素材，如您进行了改善，麻烦联系我，共同学习
;更新于：2018-06-16-20-36
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;~ and or not对应&& || !				一定要用定时器  用 TimeSetEvent 吧
;~ #UseHook							;提示，要监控所有按键操作的话，可以在脚本头加上这三行，然后双击autohotkey图标，
;~ #InstallKeybdHook				;按Ctrl+K查看所有按键，还会有VK、SC值以及作用窗口：

Process, Priority, , High			;脚本高优先级
#NoTrayIcon 						;隐藏托盘图标
#NoEnv								;不检查空变量是否为环境变量
#Persistent						;让脚本持久运行(关闭或ExitApp)
#SingleInstance Force				;跳过对话框并自动替换旧实例
#WinActivateForce					;强制激活窗口
#MaxHotkeysPerInterval 200		;时间内按热键最大次数
#HotkeyModifierTimeout 100 		;按住modifier后（不用释放后再按一次）可隐藏多个当前激活窗口
SetBatchLines -1					;脚本全速执行
SetControlDelay -1					;控件修改命令自动延时,-1无延时，0最小延时
CoordMode Menu Window				;坐标相对活动窗口
CoordMode Mouse Screen				;鼠标坐标相对于桌面(整个屏幕)
ListLines,Off           			;不显示最近执行的脚本行
SendMode Input						;更速度和可靠方式发送键盘点击
SetTitleMatchMode 2				;窗口标题模糊匹配;RegEx正则匹配
DetectHiddenWindows On				;显示隐藏窗口
SetWorkingDir %A_ScriptDir%		;当前脚本的绝对路径.不包含最后的反斜线：D:\zxh\QuickZ\Apps，通用性不如A_workingdir
;~ SetWinDelay 10					;设置在每次执行窗口命令(例如 WinActivate)后自动的延时
;~ SetKeyDelay -1					;设置每次Send和ControlSend发送键击后自动的延时,使用-1表示无延时
#InstallKeybdHook
#InstallMouseHook
;~ #KeyHistory 100
;~ #UseHook
;~ #UseHook Off

;~ Menu, Tray, Icon, %A_ScriptDir%\Word\WordZ.ico

;~ return


AnyKeyPressedOtherThanSpace(mode = "P") {
    keys = 1234567890-=qwertyuiop[]\asdfghjkl;'zxcvbnm,./
    Loop, Parse, keys
    {        
        isDown :=  GetKeyState(A_LoopField, mode)
        if(isDown)
            return True
    }   
    return False
}
;Sendinput  % uStr("原值")		;原值如果不是变量，要加双引号
uStr(str){
    charList:=StrSplit(str)
	SetFormat, integer, hex
    for key,val in charList
    out.="{U+ " . ord(val) . "}"
	return out
}


#IfWinActive ahk_exe WINWORD.EXE	;ahk_class OpusApp			;——————————————————————————————————

	;~ Selection.InsertBreak Type:=wdSectionBreakNextPage	;原始VBA
	;~ word.Selection.InsertBreak(Type:=2)					;转换后的AHK

	;~ word := ComObjActive("Word.Application")
	;~ word.Selection.Range.ListFormat.ApplyListTemplateWithLevel(word.ListGalleries(3).ListTemplates(5), 0)
	;~ Sleep 200
	;~ Send {Esc}
^+s::sendinput, {F12}

;~ ^+c::sendinput, {F2}^+{Home}^c{Esc}
;~ ^+s::sendinput, {F12}
^!Enter::	;插入分节符	word.Selection.InsertBreak(Type:=2)
	Word_insertBreak(2)
	return
^+d::
	sendinput ^c
	sendinput ^v
	sendinput ^v
	Sleep 500
	sendinput {Ctrl Up}{Shift Up}
return
!+d::
	;~ global word
	word:=ComObjActive("Word.Application")
	word.Selection.InsertDateTime(DateTimeFormat:="yyyy'年'M'月'd'日'")
return
^e::
	;~ Sleep 600
	;~ Send, {F10}fea
	wd:=ComObjActive("word.application")
	wdpdfname:=RegExReplace(ComObjActive("word.application").activedocument.fullname,"\..*$") ;. "pdf"
	wd.activedocument.ExportAsFixedFormat(wdpdfname, 17) ;~ wd.activedocument.SaveAs2("c:\acb.pdf", 17)
	return
^>#v::Word_PastePlainText()
>#tab::
	SendInput ^{F6}
	PostMessage, 0x112, 0xF020,,, A,
return

c::
	if wordformatcopy = 1
	{
		SPI_SETCURSORS := 0x57
		DllCall("SystemParametersInfo", "UInt", SPI_SETCURSORS, "UInt", 0, "UInt", 0, "UInt", 0)
		wordformatcopy := 0
	}
	else
		sendinput, c
	return
v::
	if wordformatcopy = 1
	{
		Send, {Home}^+{Down}
		Word_PasteFormat()
	}
	else
		sendinput, v
	return
Space Up::
    global space_up := true
    Send, {F18}
	sendinput {Space Up}
    return
Space::
    if AnyKeyPressedOtherThanSpace(){
        SendInput, {Blind}{Space}
        Return
    }
    space_up := False
    inputed := False
	wordformatcopy := 0
    input, UserInput, L1 T0.05, {F18}
    if (space_up){
        Send, {Blind}{Space}
        return
    }else if (StrLen(UserInput) == 1){
        Send, {Space}%UserInput%
        return
    }
    while true{
        input, UserInput, L1, {LControl}{RControl}{LAlt}{RAlt}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{Capslock}{Numlock}{PrintScreen}{Pause}{Tab}{F18}{F23}{RButton}
        ;~ MsgBox %ErrorLevel%
        if (space_up){
            if (!inputed){
                Send, {Blind}{Space}
            }
            break
			return
		}else if (UserInput == "``"){
			sendinput ^+f
			return
        }else if (ErrorLevel="EndKey:Tab"){
			;~ sendinput !op
			Send, {AppsKey}
			Send, p
			return
        }else if (ErrorLevel="EndKey:LWin"){
			Word_Dialogs_PageSetup()
			return
        }else if (ErrorLevel="EndKey:RWin"){
			sendinput, {Space}
        }else if (ErrorLevel="EndKey:F1"){
			Word_setAlignment("L")
			Word_setFontName("宋体")
			Word_setColor("ra")
			Word_SetColor("ta")
			return
        }else if (ErrorLevel="EndKey:F2"){
			Word_setAlignment("L")
			Word_setFontName("仿宋")
			Word_setColor("ra")
			Word_SetColor("ta")
			return
        }else if (ErrorLevel="EndKey:F3"){
			sendinput !p
			return
        }else if (ErrorLevel="EndKey:F4"){
			sendinput ^+!n
			return
        }else if (ErrorLevel="EndKey:F5"){
			sendinput ^!r
			return
        }else if (StrLen(UserInput) == 1) {
            inputed := True
            ;~ StringLower, UserInput, UserInput
            if (UserInput == "w")
                ;~ Send, {Up}
				Word_setFontGrow()
            else if (UserInput == "s")
                ;~ Send, {Down}
				Word_setFontShrink()
            else if (UserInput == "a")
                ;~ Send, {Left}
                Send, ^+a
            else if (UserInput == "d")
                ;~ Send, {Right}
                Send, ^+w
			else if (UserInput == "W")
                Send, {Up}
            else if (UserInput == "S")
                Send, {Down}
            else if (UserInput == "A")
                Send, {Home}
            else if (UserInput == "D")
                Send, {End}
            else if (UserInput == "1")
                ;~ Word_setAlignment("L")
				Send, ^l
            else if (UserInput == "2")
                ;~ Word_setAlignment("C")
				Send, ^e
            else  if (UserInput == "3")
                ;~ Word_setAlignment("R")
				Send, ^r
            else if (UserInput == "4")
                Word_setlinespaceShrink()
            else if (UserInput == "5")
                Word_setlinespaceGrow()
            else if (UserInput == "6")
                sendinput % uStr("^")
            else if (UserInput == ","){
                ;~ Word_setFontShrink()
				Send, ^{F6}
            }else if (UserInput == "."){
                ;~ Word_setFontGrow()
				Send, ^+{F6}
            }else if (UserInput == "6"){
				if (taggg:=!taggg){
					Word_setFirstLineIndent(0.7)
				}
				else{
					Word_setFirstLineIndent(0)
				}
            }else if (UserInput == "b"){	;插入分节符
				Word_insertBreak(2)
				;~ return
            }else if (UserInput == "c"){	;按Space+c复制文字格式，按v刷格式，没选文字或按ESC则取消格式刷。按Space+v是单次格式刷的功能
				OutputVar := strlen(ComObjActive("word.application").selection.text)
				if OutputVar=0
					Send, {Home}^+{Down}   ;^{Up}
				Word_CopyFormat()
				wordformatcopy := 1
				Sleep 150
				OCR_IBEAM = 32513
				hbeam := DllCall("LoadCursorFromFile","Str", A_ScriptDir "\Word\zgs.cur")
				DllCall( "SetSystemCursor", Uint,hbeam, Int,OCR_IBEAM )
				sendLevel 1
				sendinput {Ins}
				return
            }else if (UserInput == "C"){
                Send, {Blind}{DEL}
			}else if (UserInput == "e"){
				Word_setFontSize_eee()
				;~ sendLevel 1
				;~ sendinput {Ins}
				;~ Sleep 300
				;~ sendinput e
				return
            }else if (UserInput == "E"){
				return
            }else if (UserInput == "f"){
				Word_setTypeface_fff()
				return
            }else if (UserInput == "F"){
				Word_SetNNN()
				return
            }else if (UserInput == "g"){
				Word_setLine_ggg()
				return
            }else if (UserInput == "G"){
				sendinput ^{End}
				;~ return
            }else if (UserInput == "h"){
				Word_setShape_WrapFormat_hhh()
				return
            }else if (UserInput == "H"){
				return
            }else if (UserInput == "i"){
				word_Dialogs_Insert_Picture()
				Sleep 2000
				sendinput {Esc}
				return
            }else if (UserInput == "j"){
				Send, {PGDN}
				;~ return
            }else if (UserInput == "k"){
				Send, {PGUP}
				;~ return
            }else if (UserInput == "l"){
				Word_setFontSize(22)
				Word_setFontName("宋体")
				Word_setAlignment("C")
				Word_setBold()
				Word_setLineSpacing(1.5)
				Word_setParagraphs_Space(0,0,0)
				return
            }else if (UserInput == "m"){
				Word_insertBreak(0)
				;~ return
            }else if (UserInput == "n"){
				Word_Dialogs_Insert_Table()
				;~ return
            }else if (UserInput == "o"){
				Word_setFontSize(12)
				Word_setFontName("宋体")
				Word_setAlignment("L")
				Word_setLineSpacing(1.5)
				Word_setParagraphs_Space(0,0,0)
				Word_setFirstLineIndent(0.35)
				return
            }else if (UserInput == "p"){
				sendinput ^p
				return
            }else if (UserInput == "q"){
				Word_setformat_qqq()
				return
            }else if (UserInput == "r"){
				Word_setColorMenu_rrr()
				return
            }else if (UserInput == "t"){
				Word_setColorMenu_ttt()
				return
            }else if (UserInput == "T"){
				Word_Dialogs_Insert_Table()
				return
            }else if (UserInput == "u"){
				send ^u
				;~ return
            }else if (UserInput == "v"){
				OutputVar := strlen(ComObjActive("word.application").selection.text)
				if OutputVar=0
					Send, {Home}^+{Down}   ;^{Up}
				sendinput {F10}
				Sleep 150
				sendinput 7
				Sleep 150
				SPI_SETCURSORS := 0x57
				DllCall("SystemParametersInfo", "UInt", SPI_SETCURSORS, "UInt", 0, "UInt", 0, "UInt", 0)
				;~ wordformatcopy := 0
				return
            }else if (UserInput == "x"){
				Word_setParagraphs_Style_xxx()
				return
            }else if (UserInput == "y"){
				Word_setMargin("1,1,1,1,0.5,0.5|2,2,2,2,1.5,1.5|2.54,2.54,3.17,3.17,1.5,1.75|2.5,2.5,2.5,2.5,1.5,1.5|2.3,2.3,2.3,2.3,1.6,1.6")
				return
            }else if (UserInput == "z"){
				word文档合并工具()
				return
            }else if (UserInput == "Z"){
				Word_Zoom("-10")
				;~ return
            }else if (UserInput == "["){
				Word_SetNNN()
				return
            }else if (UserInput == "]"){
				
				return
            }else if (UserInput == "/"){
				Run %a_scriptdir%\Word\Vim-Word.jpg
				return
            }else if (UserInput == "="){
				word2016 := ComObjActive("Word.Application")
				word2016.Selection.Font.Superscript:=9999998
				;~ return
            }else if (UserInput == "-"){
				word2016 := ComObjActive("Word.Application")
				word2016.Selection.Font.Subscript:=9999998
				;~ return
            }else
                Send, {Blind}%UserInput%
        }
    }
return

#IfWinActive

word文档合并工具(){
try 
{
	wd:=ComObjActive("word.Application")
}catch e{  ;用于捕获错误,未启动word就抛出！!!
	MsgBox, 0, Title, 当前未启动word`,请先打开word`,注意不是WPS!!!, 3
	return
}
SetWorkingDir %A_ScriptDir%
Gui word文档合并工具:Font, s13
Gui word文档合并工具:Add, Text, x170 y1 w538 h50 +0x200, word文档合并工具	;(by 张磊)	;后面可修饰字体颜色,字号等;
Gui word文档合并工具:Add, Button, x420 y10 w108 h30 g关于, 关于...  ;后面可修饰字体颜色,字号等;
Gui word文档合并工具:Add, Checkbox, x40 y175 w230 h20  v子文件夹 checked, 合并子文件夹中的所有文档
Gui word文档合并工具:Add, Checkbox, x300 y175 w200 h20  v扩展名, 显示扩展名
Gui word文档合并工具:Add, CheckBox, x302 y211 w120 h23  v文件名, 插入文件名
Gui word文档合并工具:Add, CheckBox, x305 y245 w120 h23  v分节符, 插入分节符
Gui word文档合并工具:Add, ComboBox , x433 y245 w105 Choose5  v选择 g提示, 2-分节符|3-连续|4-偶数|5-奇数|7-分页符|8-分栏符|11-手动换行
Gui word文档合并工具:Add, Button, x40 y205 w110 h70  g开始合并, 开始合并
Gui word文档合并工具:Add, Button, x170 y205 w110 h70  g插入当前文档, 插入当前`n文档
Gui word文档合并工具:Add, Button, x40 y295 w110 h70  g批量查找替换, 批量查找`n替换
Gui word文档合并工具:Font
Gui word文档合并工具:Font, s14
Gui word文档合并工具:Add, Text, x28 y44 w538 h50 +0x200, 将要合并的文件夹拖入，或者定位文件夹(必须先打开word)
Gui word文档合并工具:Font
Gui word文档合并工具:Add, Edit, x32 y112 w366 h49 vEdit1
Gui word文档合并工具:Add, Button, x407 y118 w75 h47 g定位, 定位
Gui word文档合并工具:Show, w550 h380, Word文档合并工具
;~ Control, ChooseString, 7-分页符, ComboBox1,A ;复选框预先选定的方法1
;~ Control, Choose,1, ComboBox1,A ;复选框预先选定的方法2
Return
}
;autogui如何反相生成gui？

;~ word文档合并工具GuiEscape:
;~ word文档合并工具GuiClose:
    ;~ ExitApp			;————————————————————————

; End of the GUI section
关于:
MsgBox, 0, Title, 联系QQ：10000, 5
return

提示:
gui,submit,NoHide
ToolTip,将在文档中插入%选择%!
sleep 2000
ToolTip
return

开始合并: ;待合并入excel部分;
Gui, Submit
;~ MsgBox % 文件名 "-----" 扩展名  "-----" 选择
if(Edit1="")
	ExitApp			;————————————————————————
doc0:=wd.documents.add
loop ,parse,Edit1,`n,`r
{
	if (FileExist(A_LoopField)="D") ;目录判断
	{
Loop  ,%A_LoopField%\*.doc*, 0, %子文件夹% ;第三个参数:0-仅文件;1-文件+文件夹;2-仅文件夹,但是若是前面仅仅给出个母文件夹的话,可以但限制了具体文件后缀的话2就无效了;最后一个为1时为递归;
{
	st:=doc0.range.end-1
	if(文件名=1)
	{
		if(扩展名=1)
			doc0.range.InsertAfter(A_LoopFileName "`r`n")  ;带扩展名
		else
			doc0.range.InsertAfter(RegExReplace(A_LoopFileName,"`ami)\..*$") "`r`n")  ;不带扩展名
		
	doc0.Range(st,doc0.range.end-1).Style := ("标题 2")
	doc0.Range(st,doc0.range.end-1).Font.Color:=255 ;0x0000FF ;RGB(255, 0, 0)
	}
	doc0.range(doc0.range.end-1,doc0.range.end-1).Insertfile(A_LoopFileLongPath) ;插入文件
	if(分节符=1)
		doc0.range(doc0.range.end-1,doc0.range.end-1).InsertBreak(StrSplit(选择,"-")[1]*1) ;此处乘以1用于转换字符串为数字格式;
		
}
}
else if(FileExist(A_LoopField) && RegExMatch(A_LoopField,"doc"))
{

	st:=doc0.range.end-1
	if(文件名=1)
	{
		file_name:=RegExReplace(A_LoopField,"^.*\\")
		if(扩展名=1)
			doc0.range.InsertAfter(RegExReplace(A_LoopField,"`ami)^.*\\") "`r`n")  ;带扩展名
		else
			doc0.range.InsertAfter(RegExReplace(A_LoopField,"`ami)^.*\\|\..*$") "`r`n")  ;不带扩展名
	doc0.Range(st,doc0.range.end-1).Style := ("标题 2")
	doc0.Range(st,doc0.range.end-1).Font.Color:=255 ;0x0000FF ;RGB(255, 0, 0)
	}
	doc0.range(doc0.range.end-1,doc0.range.end-1).Insertfile(A_LoopField) ;插入文件
	if(分节符=1)
		doc0.range(doc0.range.end-1,doc0.range.end-1).InsertBreak(StrSplit(选择,"-")[1]*1) ;此处乘以1用于转换字符串为数字格式;	
}
}
WinActivate, % doc0.name  ;激活文档
MsgBox, 0, Title, 恭喜`,已合并完成!!!, 0.46
;;~ doc0.saveas "c:\tesd.doc"
ExitApp			;————————————————————————
return

插入当前文档:
Gui, Submit
if(Edit1="")
	ExitApp			;————————————————————————
wd:=ComObjActive("word.application")
doc0:=wd.ActiveDocument
loop ,parse,Edit1,`n,`r
{
	if (FileExist(A_LoopField)="D") ;目录判断
	{
Loop  ,%A_LoopField%\*.doc*, 0, %子文件夹% ;第三个参数:0-仅文件;1-文件+文件夹;2-仅文件夹,但是若是前面仅仅给出个母文件夹的话,可以但限制了具体文件后缀的话2就无效了;最后一个为1时为递归;
{
	st:=wd.selection.end-1
	if(文件名=1)
	{
		if(扩展名=1)
			wd.selection.InsertAfter(A_LoopFileName "`r`n")  ;带扩展名
		else
			wd.selection.InsertAfter(RegExReplace(A_LoopFileName,"`ami)\..*$") "`r`n")  ;不带扩展名
		
	doc0.Range(st,wd.selection.end-1).Style := ("标题 2")
	doc0.Range(st,wd.selection.end-1).Font.Color:=255 ;0x0000FF ;RGB(255, 0, 0)
	}
	doc0.range(wd.selection.end-1,wd.selection.end-1).Insertfile(A_LoopFileLongPath) ;插入文件
	if(分节符=1)
		doc0.range(wd.selection.end-1,wd.selection.end-1).InsertBreak(StrSplit(选择,"-")[1]*1) ;此处乘以1用于转换字符串为数字格式;
		
}
}
else if(FileExist(A_LoopField) && RegExMatch(A_LoopField,"doc"))
{

	st:=wd.selection.end-1
	if(文件名=1)
	{
		file_name:=RegExReplace(A_LoopField,"^.*\\")
		if(扩展名=1)
			doc0.range.InsertAfter(RegExReplace(A_LoopField,"`ami)^.*\\") "`r`n")  ;带扩展名
		else
			wd.selection.InsertAfter(RegExReplace(A_LoopField,"`ami)^.*\\|\..*$") "`r`n")  ;不带扩展名
	doc0.Range(st,wd.selection.end-1).Style := ("标题 2")
	doc0.Range(st,wd.selection.end-1).Font.Color:=255 ;0x0000FF ;RGB(255, 0, 0)
	}
	doc0.range(wd.selection.end-1,wd.selection.end-1).Insertfile(A_LoopField) ;插入文件
	if(分节符=1)
		doc0.range(wd.selection.end-1,wd.selection.end-1).InsertBreak(StrSplit(选择,"-")[1]*1) ;此处乘以1用于转换字符串为数字格式;	
}
}
WinActivate, % doc0.name  ;激活文档
MsgBox, 0, Title, 恭喜`,已插入完成!!!, 0.46
;~ doc0.saveas "c:\tesd.doc"
ExitApp			;————————————————————————
return


定位:
FileSelectFolder,fod,,3,文件夹选择 ;FileSelectFolder, OutputVar, ::{20d04fe0-3aea-1069-a2d8-08002b30309d}  ; 我的电脑.
if(fod="")
	return
GuiControl,, Edit1, %fod%
return



批量查找替换:
	Run, %a_scriptdir%\Word\批量查找替换(支持页眉页脚空行).docm
return


word文档合并工具GuiDropFiles:  ; 对拖放提供支持.经典代码★★★★★★★★★★★★※※※※※※
SelectedFileName := A_GuiEvent
;获取鼠标下面的控件★★★★★★★★★★★★★★★★★★★★★
MouseGetPos, , , id, control
;~ WinGetTitle, title, ahk_id %id%
WinGetClass, class, ahk_id %id%
;~ ToolTip, ahk_id %id%`nahk_class %class%`n%title%`nControl: %control%
if (control="Edit1")
{
	GuiControl,, Edit1, %SelectedFileName%  ; 在控件中显示文本.
}
if (control="Edit2")
{
	GuiControl,, Edit2, %SelectedFileName%  ; 在控件中显示文本.
}
return


#Include %a_scriptdir%\Word\VIMD_Word.ahk