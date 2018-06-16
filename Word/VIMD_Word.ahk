    global word

/* 
;Word_VIMD批处理使用示例
;不想设置的内容，直接在行前添加【;】，注释当前行即可
 ---------------------------------------------------------------------------------------
;设置字号
;42=初号，36=小初，26=一号，24=小一，22=二号，18=小二,16=三号，15=小三，14=四号，12=小四，10=五号，9=小五,7=六号，6=小六，5=七号，5=八号
Word_setFontSize(10)

;设置字体
;宋体，黑体，仿宋，微软雅黑
Word_setFontName("宋体")

;设置字体、高亮颜色
;高亮颜色，ra=自动，rb=蓝色，rB=浅蓝，rg=绿色，rG=浅绿，r1=灰1，r2=灰2，r3=灰3，r4=灰4，rk=黑色，rn=无，ro=橙色，rO=浅橙，rr=红色，rR=浅红，rw=白色，ry=黄色，rY=浅黄
Word_setColor("ry")
;字体颜色，ta=自动，tb=蓝色，tB=浅蓝，tg=绿色，tG=浅绿，t1=灰1，t2=灰2，t3=灰3，t4=灰4，tk=黑色，to=橙色，tO=浅橙，tr=红色，tR=浅红，tw=白色，ty=黄色，tY=浅黄
Word_setColor("tr")

;设置对齐
;对齐格式，L=左对齐，C=居中，R=右对齐，J=两端对齐，D=分散对齐
Word_setAlignment("L")

;设置粗体
Word_setBold()
;设置斜体
Word_setItalic()
;设置下划线
Word_setUnderline()
;设置删除线
Word_setStrikeThrough()

;设置行间距
;1,5=1.5倍行距
Word_setLineSpacing(1.5)
;设置段前段后0行，（段前行数，段后行数，1=自动|0=0行）
Word_setParagraphs_Space(0,0,0)

;设置首行缩进
;两字符=0.35
Word_setFirstLineIndent(0.35)

;设置页边距
;(上边距，下边距，左边距，右边距，页眉边距，页脚边距）
Word_PageSetup_Margin(1,1,1,1,0.5,0.5)

;设置纸张方面
;自动转换
Word_PageSetup_Orientation()

*/


Word_setFontGrowzzz(){
	word := ComObjActive("Word.Application")
	word.selection.font.Grow
}

;~ {
;~ ^p::
	;~ Word_setMargin("1,1,1,1,0.5,0.5|2,2,2,2,1.5,1.5|3.17,3.17,2.54,2.54,1.5,1.75")
;~ return
;~ }

;===================================================================
/*
    函数:  Word_Get
    作用: 返回word对象
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_Get() {
	global word
	word:=ComObjActive("Word.Application")
return
}

/*
    函数:  Word_Destroy
    作用: 摧毁word对象
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_Destroy(){
	global word
	word:=
}

/*
    函数:  Word_OpenFile
    作用:  打开对话框
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_OpenFile(){
	sendinput,^o
}

/*
    函数:  Word_PasteFormat
    作用:  粘贴格式
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_PasteFormat(){
	send,^+v
}

/*
    函数:  Word_CopyFormat
    作用:  复制格式
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_CopyFormat(){
	send,^+c
}

/*
    函数:  Word_SaveFile
    作用:  保存
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_SaveFile(){
	sendinput,^s
}

/*
    函数:  Word_SaveFileAs
    作用:  另存为
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_SaveFileAs(){
	;~ sendinput,^+s
	sendinput, {F12}
}

/*
    函数:  Word_CloseWord
    作用:  关闭
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_CloseWord(){
	global word
	Word_Get()
	try
		word.ActiveDocument.Close
	Word_Destroy()
	
}

/*
    函数:  Word_SaveAndExit
    作用:  保存退出
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
	wdDialogFileSummaryInfo:=86
*/
Word_SaveAndExit(){
	global word
	Word_Get()
	Word_Dialogs(86)
	if !(word.ActiveDocument.Saved)
		send,^s
	word.ActiveDocument.Close
	Word_Destroy()
	
}


/*
    函数:  Word_setFontSize
    作用: 设置字号
    参数: fSize
			42=初号，36=小初，26=一号，24=小一，22=二号，18=小二
			16=三号，15=小三，14=四号，12=小四，10=五号，9=小五
			7=六号，6=小六，5=七号，5=八号
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFontSize(fSize:=10){
	if (fSize="")
		return
	global word
	Word_Get()
	try
		word.selection.font.size:=fSize
	Word_Destroy()
	
}

/*
    函数:  Word_setFontSize
    作用: 设置字号
    参数: fSize
			42=初号，36=小初，26=一号，24=小一，22=二号，18=小二
			16=三号，15=小三，14=四号，12=小四，10=五号，9=小五
			7=六号，6=小六，5=七号，5=八号
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFontSizee(fSize:=""){
	global VimD
	global word
	Word_Get()
	if (fSize=""){
		MyType:=substr(VimD.HotKeyStr,1,1)
		MyIndex:=substr(VimD.HotKeyStr,2)
	} else {
		MyType:=substr(fSize,1,1)
		MyIndex:=substr(fSize,2)
	}
	if ((MyIndex+0)>0)
		MyIndex=g%MyIndex%
	;42=初号，36=小初，26=一号，24=小一，22=二号，18=小二，16=三号，15=小三
	;14=四号，12=小四，10=五号，9=小五，7=六号，6=小六，5=七号，5=八号
	MyHighColorArray1:={"a":42,"s":36,"g1":26,"q":24,"g2":22,"w":18,"g3":16,"e":15,"g4":14,"r":12,"g5":10,"t":9,"g6":7,"y":6,"g7":5}
	MyHighColorArray2:={"R":5,"B":3,"Y":14,"G":11,"O":14}
	if (MyType="e")
	{
	if RegExMatch(MyIndex,"^[a-z0-9]+$")
		MyValue:=MyHighColorArray1[MyIndex]
	else
		MyValue:=MyHighColorArray2[MyIndex]
	try
		word.selection.font.size:=MyValue
	}
	Word_Destroy()

}

Word_setFontSize_eee(){
	QuickInputList=
	(Ltrim
		&a)42文字_初号
		&s)36文字_小初
		&1)26文字_一号
		&q)24文字_小一
		&2)22文字_二号
		&w)18文字_小二
		&3)16文字_三号
		&e)15文字_小三
		&4)14文字_四号
		&r)12文字_小四
		&5)10文字_五号
		&t)09文字_小五
		&6)07文字_六号
		&y)06文字_小六
		&7)05文字_七号
	)

	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlere
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlere:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		outputStr:="e" . outputStr
		Word_setFontSizee(outputStr)
	return
}


/*
    函数:  Word_setformat
    作用: 设置格式
    参数: qformat
    返回:
    作者:  qzSpace
    版本:  0.1
*/
Word_setformat(qformat:=""){
	global VimD
	global word
	Word_Get()
	if (qformat=""){
		MyType:=substr(VimD.HotKeyStr,1,1)
		MyIndex:=substr(VimD.HotKeyStr,2)
	} else {
		MyType:=substr(qformat,1,1)
		MyIndex:=substr(qformat,2)
	}	
	;~ MsgBox %MyType%_____%MyIndex%
	;~ if ((MyIndex+0)>0)
		;~ MyIndex=g%MyIndex%
	;~ MyHighColorArray1:={"a":42,"s":36,"g1":26,"q":24,"g2":22,"w":18,"g3":16,"e":15,"g4":14,"r":12,"g5":10,"t":9,"g6":7,"y":6,"g7":5}
	;~ MyHighColorArray2:={"R":5,"B":3,"Y":14,"G":11,"O":14}
	if (MyType="q")
	{
	;~ if RegExMatch(MyIndex,"^[a-z0-9]+$")
	if (MyIndex="1")
		Word_setAlignment("L")
	else if (MyIndex="2")
		Word_setAlignment("C")
	else if (MyIndex="3")
		Word_setAlignment("R")
	else if (MyIndex="4")
		Word_setAlignment("J")
	else if (MyIndex="5")
		Word_setAlignment("D")
	else if (MyIndex="b")
		Word_setBold()
	else if (MyIndex="d")
		Word_setStrikeThrough()
	else if (MyIndex="i")
		Word_setItalic()
	else if (MyIndex="u")
		Word_setUnderline()
	}
	Word_Destroy()

}

Word_setformat_qqq(){
	QuickInputList=
	(Ltrim
		&1)Word_左对齐
		&2)Word_居中
		&3)Word_右对齐
		&4)Word_两站对齐
		&5)Word_分散对齐
		&b)Word_设置or取消粗体
		&d)Word_设置or取消删除线
		&i)Word_设置or取消斜体
		&u)Word_设置or取消下划线
	)
	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerq
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerq:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		outputStr:="q" . outputStr
		Word_setformat(outputStr)
	return
}

/*
    函数:  Word_setlinespaceGrow
    作用: 设置行号加大
    参数:
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setlinespaceGrow(){
	global word
	;~ word := ComObjActive("Word.Application")
	Word_Get()
	rules := word.Selection.ParagraphFormat.LineSpacing
	rules := rules + 2
	try
		word.selection.ParagraphFormat.LineSpacing := rules
	Word_Destroy()

}

/*
    函数:  Word_setlinespaceShrink
    作用: 设置行号减小
    参数:
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setlinespaceShrink(){
	global word
	;~ word := ComObjActive("Word.Application")
	Word_Get()
	rules := word.Selection.ParagraphFormat.LineSpacing
	rules := rules - 2
	try
		word.selection.ParagraphFormat.LineSpacing := rules
	Word_Destroy()

}

/*
    函数:  Word_setFontGrow
    作用: 设置字号加大
    参数:
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFontGrow(){
	global word
	Word_Get()
	try
		word.selection.font.Grow
	Word_Destroy()

}

/*
    函数:  Word_setFontShrink
    作用: 设置字号减小
    参数:
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFontShrink(){
	global word
	Word_Get()
	try
		word.selection.font.Shrink
	Word_Destroy()

}

/*
    函数:  Word_setFontColor
    作用: 设置字体颜色
    参数: fMyColor,如果为空，则自动获取热键值，否则设置
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setColor(fMyColor:=""){
	global VimD
	global word
	Word_Get()
	if (fMyColor=""){
		MyType:=substr(VimD.HotKeyStr,1,1)
		MyIndex:=substr(VimD.HotKeyStr,2)
	} else {
		MyType:=substr(fMyColor,1,1)
		MyIndex:=substr(fMyColor,2)
	}
	if ((MyIndex+0)>0)
		MyIndex=g%MyIndex%
	;红色=6，蓝色=2，黄色=7，绿色=4，橙色=14，灰1=16，灰2=16，灰3=15，灰4=9，白色=8，黑色=1，自动=0，无=0
	;浅红=5，浅蓝=3，浅黄=14，浅绿=11，浅橙=
	MyHighColorArray1:={"r":6,"b":2,"y":7,"g":4,"o":14,"g1":16,"g2":16,"g3":15,"g4":9,"w":8,"k":1,"a":0,"n":0}
	MyHighColorArray2:={"R":5,"B":3,"Y":14,"G":11,"O":14}
	if (MyType="r")		;高亮颜色
	{
		if RegExMatch(MyIndex,"^[a-z0-9]+$")
			MyValue:=MyHighColorArray1[MyIndex]
		else
			MyValue:=MyHighColorArray2[MyIndex]
		try
			word.Options.DefaultHighlightColorIndex:=MyValue
			word.selection.Range.HighlightColorIndex:=MyValue
	}
	;红色=255，蓝色=16711680，黄色=65535，绿色=32768，橙色=26367，灰1=142770819，灰2=12632256，灰3=8421504，灰4=2500134，白色=16777215，黑色=0，自动=-16777216，无=0
	;浅红=8420607，浅蓝=16737843，浅黄=10092543，浅绿=13434828，浅橙=39423
	MyFontColorArray1:={"r":255,"b":16711680,"y":65535,"g":32768,"o":26367,"g1":142770819,"g2":12632256,"g3":8421504,"g4":2500134,"w":16777215,"k":0,"a":-16777216,"n":0}
	MyFontColorArray2:={"R":8420607,"B":16737843,"Y":10092543,"G":13434828,"O":39423}
	if (MyType="t")		;字体颜色
	{
		if RegExMatch(MyIndex,"^[a-z0-9]+$")
			MyValue:=MyFontColorArray1[MyIndex]
		else
			MyValue:=MyFontColorArray2[MyIndex]
		try
			word.selection.Font.Color:=MyValue
	}
	Word_Destroy()
	;~ gosub 597EE6A6-A3FA-4E00-8A3C-5AA212D6D471

}

Word_setColorMenu_ttt(){
	QuickInputList=
	(Ltrim
		&1)文字颜色_无
		&2)文字颜色_浅灰
		&3)文字颜色_中灰
		&4)文字颜色_深灰
		&a)文字颜色_无色
		&n)文字颜色_无色
		&b)文字浅蓝   B_深蓝
		&g)文字浅绿   G_深绿
		&o)文字浅橘   O_深橘
		&r)文字浅红   R_深红
		&y)文字浅黄   Y_深黄
		&k)文字颜色_无
		&w)文字颜色_白色
	)

	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlert
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlert:
		outputStr:=substr(A_ThisMenuItem,2,1)
		;【Kawvin】2018-04-08-18:15:17
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		outputStr:="t" . outputStr
		Word_setColor(outputStr)
	return
}
Word_setColorMenu_rrr(){
	QuickInputList=
	(Ltrim
		&1)底色颜色_无
		&2)底色颜色_浅灰
		&3)底色颜色_中灰
		&4)底色颜色_深灰
		&a)底色颜色_无色
		&n)底色颜色_无色
		&b)底色浅蓝   B_深蓝
		&g)底色浅绿   G_深绿
		&o)底色浅橘   O_深橘
		&r)底色浅红   R_深红
		&y)底色浅黄   Y_深黄
		&k)底色颜色_无
		&w)底色颜色_白色
	)

	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerr
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerr:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr
		outputStr:="r" . outputStr
		Word_setColor(outputStr)
	return
}


/*
    函数:  Word_setAlignment
    作用: 设置对齐
    参数: fAlignment	L=Left，C=Center，R=Right，J=Justify，D=Distribute，B=JustifyMed
		;枚举常量类型：WdParagraphAlignment
		wdAlignParagraphCenter:=1
		wdAlignParagraphDistribute:=4
		wdAlignParagraphJustify:=3
		wdAlignParagraphJustifyHi:=7
		wdAlignParagraphJustifyLow:=8
		wdAlignParagraphJustifyMed:=5
		wdAlignParagraphLeft:=0
		wdAlignParagraphRight:=2
		wdAlignParagraphThaiJustify:=9
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setAlignment(fAlignment:="L"){
	global VimD
	if (fAlignment="")
		return
	global word
	Word_Get()
	aArray:={"L":0,"C":1,"R":2,"J":3,"D":4,"B":9}
	try
		word.Selection.ParagraphFormat.Alignment :=aArray[fAlignment]
	Word_Destroy()

}

/*
    函数:  Word_setBold
    作用: 设置粗体
    参数: wdToggle:=9999998
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setBold(){
	global word
	Word_Get()
	try
		word.selection.font.Bold:=9999998
	Word_Destroy()

}

/*
    函数:  Word_setItalic
    作用: 设置斜体
    参数: wdToggle:=9999998
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setItalic(){
	global word
	Word_Get()
	try
		word.selection.font.Italic:=9999998
	Word_Destroy()

}

/*
    函数:  Word_setUnderline
    作用: 设置下划线
    参数:
			;枚举常量类型：WdUnderline
			wdUnderlineDash:=7
			wdUnderlineDashHeavy:=23
			wdUnderlineDashLong:=39
			wdUnderlineDashLongHeavy:=55
			wdUnderlineDotDash:=9
			wdUnderlineDotDashHeavy:=25
			wdUnderlineDotDotDash:=10
			wdUnderlineDotDotDashHeavy:=26
			wdUnderlineDotted:=4
			wdUnderlineDottedHeavy:=20
			wdUnderlineDouble:=3
			wdUnderlineNone:=0
			wdUnderlineSingle:=1
			wdUnderlineThick:=6
			wdUnderlineWavy:=11
			wdUnderlineWavyDouble:=43
			wdUnderlineWavyHeavy:=27
			wdUnderlineWords:=2
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setUnderline(){
	global word
	Word_Get()
	try
	{
		if (word.selection.font.Underline<>0)
			word.selection.font.Underline:=0
		else
			word.selection.font.Underline:=1
	}
	Word_Destroy()

}

/*
    函数:  Word_setStrikeThrough
    作用: 设置删除线
    参数: wdToggle:=9999998
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setStrikeThrough(){
	global word
	Word_Get()
	try
		word.selection.font.StrikeThrough:=9999998
	Word_Destroy()

}

/*
    函数:  Word_setLineSpacing
    作用: 设置行距
    参数: fLineSpacing
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setLineSpacing(fLineSpacing:=1){
	if (fLineSpacing="")
		return
	global word
	Word_Get()
	fLinesPoints:=fLineSpacing*12
	try
		word.selection.ParagraphFormat.LineSpacing:=fLinesPoints
	Word_Destroy()

}


/*
    函数:  Word_setLineSpacing
    作用: 设置字号
    参数: fSize
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setLine(fSize:=""){			;gggggg
	global VimD
	global word
	Word_Get()
	if (fSize=""){
		MyType:=substr(VimD.HotKeyStr,1,1)
		MyIndex:=substr(VimD.HotKeyStr,2)
	} else {
		MyType:=substr(fSize,1,1)
		MyIndex:=substr(fSize,2)
	}
	;~ if ((MyIndex+0)>0)
		;~ MyIndex=f%MyIndex%
	;42=初号，36=小初，26=一号，24=小一，22=二号，18=小二，16=三号，15=小三
	;14=四号，12=小四，10=五号，9=小五，7=六号，6=小六，5=七号，5=八号
	;~ MyHighColorArray1:={"f1":1,"f2":2,"f3":3,"f4":4,"f5":5,"a":18,"g3":16,"e":15,"g4":14,"r":12,"g5":10,"t":9,"g6":7,"y":6,"g7":5}
	;~ MyHighColorArray2:={"R":5,"B":3,"Y":14,"G":11,"O":14}
	if (MyType="g")
	{
		if RegExMatch(MyIndex,"^[1-9]+$"){
			MyIndex:=MyIndex*12
			try
				word.selection.ParagraphFormat.LineSpacing:=MyIndex
		}else if (MyIndex="0"){
			Word_setParagraphs_Space(0,0,0)
		}else if (MyIndex="a"){
			Word_setParagraphs_Space("","",1)
		}else if (MyIndex="f"){
			sendinput ^+f
		}else if (MyIndex="g"){			
			sendinput !op
		}else if (MyIndex="s"){
			Word_setFirstLineIndent(0.35)
		}else if (MyIndex="d"){
			Word_setFirstLineIndent(0.7)
		}
	}
	Word_Destroy()
}

Word_setLine_ggg(){
	QuickInputList=
	(Ltrim
		&0)段前间距_0行
		&1)段落行距_1倍
		&2)段落行距_2倍
		&3)段落行距_3倍
		&4)段落行距_4倍
		&5)段落行距_5倍
		&a)段前间距_自动
		&f)字体设置
		&g)段落设置
		&s)首行缩进2字符0.35
		&d)首行缩进4字符0.7
	)
	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerg
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerg:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		outputStr:="g" . outputStr
		Word_setLine(outputStr)
	return
}

/*
    函数:  Word_setFirstLineIndent
    作用: 设置首行缩进
    参数: fFirstLineIndent	，0.35=2字符
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFirstLineIndent(fFirstLineIndent:=0.35){
	if (fFirstLineIndent="")
		return
	global word
	Word_Get()
	try
		word.Selection.ParagraphFormat.FirstLineIndent := CentimetersToPoints(fFirstLineIndent)
	Word_Destroy()

}

/*
    函数:  Word_insertBreak(fType)
    作用: 插入分隔符
    参数: fType，分隔符类型
    返回:
    作者:  Kawvin
    版本:  0.1
	wdSectionBreakNextPage:=2
	wdPageBreak:=7	分页符，Ctrl+Return
	0	分页符，Ctrl+Enter
*/
Word_insertBreak(fType:=0){
	if (fType="")
		return
	global word
	Word_Get()
	try
		word.Selection.InsertBreak(fType)
	Word_Destroy()

}

/*
    函数:  Word_setFontName
    作用: 设置字体
    参数: fFontName
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setFontName(fFontName:="宋体"){
	if (fFontName="")
		return
	global word
	Word_Get()
	try
		word.Selection.Font.Name :=fFontName
	Word_Destroy()

}

/*
    函数:  Word_Zoom
    作用: 比例设置
    参数: +n，比例增加n
			 -n，比例减少n
			 n，设置比例为n
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_Zoom(fZoom:="-5"){
	if (fZoom="")
		return
	global word
	Word_Get()
	if (regexmatch(fZoom,"\+[0-9]+")) {		;放大
		ZoomNum:=substr(fZoom,2)
		try
			word.ActiveWindow.ActivePane.View.Zoom.Percentage+=ZoomNum
	}else if (regexmatch(fZoom,"\-[0-9]+")) {	;缩小
		ZoomNum:=substr(fZoom,2)
		try
			word.ActiveWindow.ActivePane.View.Zoom.Percentage-=ZoomNum
	}else if (regexmatch(fZoom,"[0-9]+")) {		;设置比例 		
		ZoomNum:=substr(fZoom,1)
		if(ZoomNum<10)
			ZoomNum:=10
		if(ZoomNum>500)
			ZoomNum:=500
		try
			word.ActiveWindow.ActivePane.View.Zoom.Percentage:=ZoomNum
	}else {
		try
			word.ActiveWindow.ActivePane.View.Zoom.Percentage-=5
	}
	Word_Destroy()
	
}

/*
    函数:  Word_PastePlainText
    作用: 粘贴为无格式文本
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
	wdFormatPlainText:=22
*/
Word_PastePlainText(){
	global word
	Word_Get()
	try
		word.Selection.PasteAndFormat(22)
	Word_Destroy()
	
}

/*
    函数:  Word_SetNNN
    作用: 各类设置
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_SetNNN(){
	prompt=
	(LTrim
	调整行高/列宽/字号，输入格式：
	1)设置字号为n，命令格式：f20
	2)设置行距为n，命令格式：l2
	3)设置显示比例为n，命令格式：z80
	)
	InputBox,cmd_Buff,请输入命令,%prompt%, , 300,200   ;``
	if ErrorLevel
		return
	if (cmd_Buff="")
		return

	if RegExMatch(cmd_Buff,"i)f(\d)+")		;设置字号
	{
		buff_string := substr(cmd_buff, 2)
		Word_setFontSize(buff_string)
	} else if RegExMatch(cmd_Buff,"i)l(\d)+")  {
		buff_string := substr(cmd_buff, 2)
		Word_setLineSpacing(buff_string)
	} else if RegExMatch(cmd_Buff,"i)z(\d)+")  {
		buff_string := substr(cmd_buff, 2)
		Word_Zoom(buff_string)
	} else   {
		
	}
	
}

/*
    函数:  Word_Dialogs
    作用: 显示对话框
    参数: fDig，对话框常数
    返回:
    作者:  Kawvin
    版本:  0.1
	wdDialogFilePrintSetup:=97
	wdDialogFilePageSetup:=178
	wdDialogFileSummaryInfo:=86
	wdDialogInsertPicture:=163
	wdDialogFilePrint:=88
	wdDialogTableInsertTable:=129


*/
Word_Dialogs(fDig){
	if (fDig="")
		return
	global word
	Word_Get()
	try
		word.Dialogs(fDig).show
	;Word_Destroy()
}

/*
    函数:  Word_Dialogs_Insert_Picture
    作用: 插入图片对话框
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_Dialogs_Insert_Picture(){
	global word
	Word_Dialogs(163)
	Word_Destroy()
	
}

/*
    函数:  Word_Dialogs_Insert_Table
    作用: 插入表格对话框
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_Dialogs_Insert_Table(){
	global word
	Word_Dialogs(129)
	Word_Destroy()
	
}

/*
    函数:  Word_Dialogs_PageSetup
    作用: 页面设置对话框
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_Dialogs_PageSetup(){
	global word
	Word_Dialogs(178)
	Word_Destroy()
	
}

/*
    函数:  Word_Dialogs_PrintPreview
    作用: 打印对话框
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_Dialogs_PrintPreview(){
	send,^p
}
	
/*
    函数:  Word_setShape_WrapFormat
    作用: 设置图片环绕方式
    参数: fWrapType		0=四周型,1=紧密型,2=穿越型环绕,31=衬于文字下方,32=浮于文字上方,4=上下型环绕,7=嵌入型
    返回:
    作者:  Kawvin
    版本:  0.1
	;枚举常量类型：WdWrapType
	wdWrapSquare:=0				四周型
	wdWrapTight:=1				紧密环绕型
	wdWrapThrough:=2			穿越型环绕
	wdWrapNone:=3				文字下方或文字上方
	wdWrapTopBottom:=4		上下型环绕
	wdWrapInline:=7				嵌入型
	
	wdSelectionInlineShape:=7
	wdSelectionShape:=8
*/
Word_setShape_WrapFormat(fWrapType:=7){
	if (fWrapType="")
		return
	global word
	Word_Get()
	if (word.Selection.type=7) {
		SelPic:=word.Selection.InlineShapes(1)
		if (fWrapType=0)||(fWrapType=1)||(fWrapType=2)||(fWrapType=31)||(fWrapType=32)
			SelPic:=SelPic.ConvertToShape()			;类型转换
		if (fWrapType=0)||(fWrapType=1)||(fWrapType=2)||(fWrapType=4)||(fWrapType=7) {
			try
				SelPic.WrapFormat.Type := fWrapType
		} else if (fWrapType=31) {
			try {
				SelPic.WrapFormat.Type :=3
				SelPic.ZOrder(5)			;图片衬于文字下方
			}
		} else if (fWrapType=32) {
			try {
				SelPic.WrapFormat.Type :=3
				SelPic.ZOrder(4)			;图片浮于文字上方
			}
		}
		Word_Destroy()
		
		return
	}
	if (word.Selection.type=8)  {
		SelPic:=word.Selection.ShapeRange(1)
		if (fWrapType=7)
			SelPic:=SelPic.ConvertToInlineShape()				;类型转换
		if (fWrapType=0)||(fWrapType=1)||(fWrapType=2)||(fWrapType=4)||(fWrapType=7) {
			try
				SelPic.WrapFormat.Type := fWrapType
		} else if (fWrapType=31) {
			try {
				SelPic.WrapFormat.Type :=3
				SelPic.ZOrder(5)			;图片衬于文字下方
			}
		} else if (fWrapType=32) {
			try {
				SelPic.WrapFormat.Type :=3
				SelPic.ZOrder(4)			;图片浮于文字上方
				
			}
		}
		Word_Destroy()
		
		return
	}
	Word_Destroy()
	
}



Word_setShape_WrapFormat_hhh(){
	QuickInputList=
	(Ltrim
		&g)图片_四周型
		&h)图片_嵌入型
		&j)图片_紧密环绕型
		&n)图片_衬于文字下方
		&t)图片_穿越型环绕
		&u)图片_上下型环绕
		&y)图片_浮于文字上方
	)
	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerh
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerh:
		outputStr:=substr(A_ThisMenuItem,2,1)
		;~ if GetKeyState("Shift")
			;~ Stringupper, outputStr, outputStr 
		;~ outputStr:="f" . outputStr
		if outputStr=g
			outputStr:=0
		if outputStr=h
			outputStr:=7
		if outputStr=j
			outputStr:=1
		if outputStr=n
			outputStr:=31
		if outputStr=t
			outputStr:=2
		if outputStr=u
			outputStr:=4
		if outputStr=y
			outputStr:=32		
		Word_setShape_WrapFormat(outputStr)
	return
}


/*
    函数:  Word_setParagraphs_Style
    作用: 设置大纲级别
    参数: fStyle		0=正文，1=1级，...9=9级
    返回:
    作者:  Kawvin
    版本:  0.1
	wdStyleNormal:=-1
	wdStyleHeading1:=-2
	wdStyleHeading2:=-3
	wdStyleHeading3:=-4
	wdStyleHeading4:=-5
	wdStyleHeading5:=-6
	wdStyleHeading6:=-7
	wdStyleHeading7:=-8
	wdStyleHeading8:=-9
	wdStyleHeading9:=-10
*/
Word_setParagraphs_Style(fStyle:=0){
	if (fStyle="")
		return
	global word
	Word_Get()
	try
		word.Selection.Range.Style := word.ActiveDocument.Styles(0-fStyle-1)
	Word_Destroy()
	
}

Word_setParagraphs_Style_xxx(){
	QuickInputList=
	(Ltrim
		&1)大纲_1级目录
		&2)大纲_2级目录
		&3)大纲_3级目录
		&4)大纲_4级目录
		&5)大纲_5级目录
		&6)大纲_6级目录
		&a)大纲_正文
		&s)大纲_目录降级
		&w)大纲_目录升级
		&z)选择_表
	)
	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerx
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerx:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		;~ outputStr:="x" . outputStr
		if outputStr=w
		{
			Word_setParagraphs_Promote()
			return
		}
		if outputStr=s
		{
			Word_setParagraphs_Demote()
			return
		}
		if outputStr=z
		{
			sendinput {F10}jlkt
			return
		}
		if outputStr=a
			outputStr:=0
		Word_setParagraphs_Style(outputStr)
	return
}

/*
    函数:  Word_Setlistlevel
    作用: 设置列表级别
    参数: eStyle
    返回:
    作者:  天甜
    版本:  0.1
*/
Word_Setlistlevel(eStyle:=0){
	if (eStyle="")
		return
	global word
	Word_Get()
	try
		word.Selection.Range.SetListLevel(Level:=eStyle)
	Word_Destroy()
	
}


/*
    函数:  Word_setTypeface
    作用: 设置字号
    参数: fSize
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setTypeface(fSize:=""){			;FFFFFF
	global VimD
	global word
	Word_Get()
	if (fSize=""){
		MyType:=substr(VimD.HotKeyStr,1,1)
		MyIndex:=substr(VimD.HotKeyStr,2)
	} else {
		MyType:=substr(fSize,1,1)
		MyIndex:=substr(fSize,2)
	}
	;~ if ((MyIndex+0)>0)
		;~ MyIndex=f%MyIndex%
	;42=初号，36=小初，26=一号，24=小一，22=二号，18=小二，16=三号，15=小三
	;14=四号，12=小四，10=五号，9=小五，7=六号，6=小六，5=七号，5=八号
	;~ MyHighColorArray1:={"f1":1,"f2":2,"f3":3,"f4":4,"f5":5,"a":18,"g3":16,"e":15,"g4":14,"r":12,"g5":10,"t":9,"g6":7,"y":6,"g7":5}
	;~ MyHighColorArray2:={"R":5,"B":3,"Y":14,"G":11,"O":14}
	if (MyType="f")
	{
		if RegExMatch(MyIndex,"^[0-9]+$"){
			try
			word.Selection.Range.SetListLevel(Level:=MyIndex)
		}else if (MyIndex="a"){
			sendinput {F10}	;Office快捷键
			Sleep 50
			sendinput hm{home}{down 3}{right 2}{enter}
			return
		}else if (MyIndex="b"){
			sendinput {F10}	;Office快捷键
			Sleep 50
			sendinput hm{home}{down 4}{right}{enter}
			return
		}else if (MyIndex="f"){
			Word_setFontName("仿宋")
		}else if (MyIndex="h"){
			Word_setFontName("黑体")
		}else if (MyIndex="q"){
			sendinput {F10}	;Office快捷键
			Sleep 50
			sendinput hm{home}{down 4}{enter}
			return
		}else if (MyIndex="s"){
			Word_setFontName("宋体")
		}else if (MyIndex="t"){
			Word_setFontName("Times New Roman")
		}else if (MyIndex="y"){
			Word_setFontName("微软雅黑")
		}
	}
	Word_Destroy()
	
}

Word_setTypeface_fff(){
	QuickInputList=
	(Ltrim
		&1)多级列表_1级
		&2)多级列表_2级
		&3)多级列表_3级
		&4)多级列表_4级
		&5)多级列表_5级
		&a)多级列表_zxh常用
		&b)列表_表格序号
		&q)列表_一二三
		&f)字体_仿宋
		&h)字体_黑体
		&s)字体_宋体
		&t)字体_TimeNewRoman
		&y)字体_微软雅黑
	)
	menu,KyMenu_QuickInput,Add
	menu,KyMenu_QuickInput,DeleteAll
	Loop,parse,QuickInputList,`n,`r
	{
		if (A_LoopField="")
			continue
		if (A_LoopField="-")
			menu,KyMenu_QuickInput,Add
		else
			menu,KyMenu_QuickInput,Add,% A_LoopField,KyMenu_QuickInput_Handlerf
	}
	menu,KyMenu_QuickInput,show
	return

	KyMenu_QuickInput_Handlerf:
		outputStr:=substr(A_ThisMenuItem,2,1)
		if GetKeyState("Shift")
			Stringupper, outputStr, outputStr 
		outputStr:="f" . outputStr
		Word_setTypeface(outputStr)
	return
}


/*
    函数:  Word_Setlistlevel_zxh_cy
    作用: 设置zxh自定义列表级别
    参数: 
    返回:
    作者:  天甜
    版本:  0.1
*/
Word_Setlistlevel_zxh_cy(){
	global word
	Word_Get()
		try
		{
		LL1:=word.ListGalleries(3).ListTemplates(1).ListLevels(1)
		LL1.NumberFormat := "%1、"
		LL1.NumberPosition := 0*28.35		;对齐位置-相当于首行缩进
		LL1.TextPosition := 0.75*28.35		;文本左缩进
		LL1.ResetOnHigher := 0
		LL1.TrailingCharacter := 2
		LL1.NumberStyle := 0
		LL1.TabPosition := 9999999
		LL1.ResetOnHigher := 0
		LL1.StartAt := 1

		LL2:=word.ListGalleries(3).ListTemplates(1).ListLevels(2)
		LL2.NumberFormat := "%1.%2、"
		LL2.NumberPosition := 0.5*28.35
		LL2.TextPosition := 1.5*28.35
		LL2.ResetOnHigher := 1
		LL2.TrailingCharacter := 2
		LL2.NumberStyle := 0
		LL2.TabPosition := 9999999
		LL2.ResetOnHigher := 0
		LL2.StartAt := 1

		LL2:=word.ListGalleries(3).ListTemplates(1).ListLevels(3)
		LL2.NumberFormat := "%1.%2.%3、"
		LL2.NumberPosition := 1*28.35
		LL2.TextPosition := 2*28.35
		LL2.ResetOnHigher := 2
		LL2.TrailingCharacter := 2
		LL2.NumberStyle := 0
		LL2.TabPosition := 9999999
		LL2.ResetOnHigher := 0
		LL2.StartAt := 1

		LL2:=word.ListGalleries(3).ListTemplates(1).ListLevels(4)
		LL2.NumberFormat := "%1.%2.%3.%4、"
		LL2.NumberPosition := 1.5*28.35
		LL2.TextPosition := 2.5*28.35
		LL2.ResetOnHigher := 3
		LL2.TrailingCharacter := 2
		LL2.NumberStyle := 0
		LL2.TabPosition := 9999999
		LL2.ResetOnHigher := 0
		LL2.StartAt := 1

		LL2:=word.ListGalleries(3).ListTemplates(1).ListLevels(5)
		LL2.NumberFormat := "%1.%2.%3.%4.%5、"
		LL2.NumberPosition := 2*28.35
		LL2.TextPosition := 3*28.35
		LL2.ResetOnHigher := 4
		LL2.TrailingCharacter := 2
		LL2.NumberStyle := 0
		LL2.TabPosition := 9999999
		LL2.ResetOnHigher := 0
		LL2.StartAt := 1

		word.ListGalleries(3).ListTemplates(1).Name:=""
		word.Selection.Range.ListFormat.ApplyListTemplateWithLevel(word.ListGalleries(3).ListTemplates(1), 0)
		}
		word.Selection.Range.SetListLevel(Level:=1)
	Word_Destroy()
	
}


/*
    函数:  Word_setParagraphs_Space
    作用: 设置段落段前、段后间距
    参数: fBefroe，段前行数，fAfter，段后行数，fAuto，自动行距(1,0)
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setParagraphs_Space(fBefroe:="",fAfter:="",fAuto:=""){
	global word
	Word_Get()
	if(fBefroe!="")
		try
			word.Selection.ParagraphFormat.SpaceBefore :=fBefroe
	if(fAfter!="")
		try
			word.Selection.ParagraphFormat.SpaceAfter :=fAfter
	if(fAuto=1){
		try{
			word.Selection.ParagraphFormat.SpaceBeforeAuto := True
			word.Selection.ParagraphFormat.SpaceAfterAuto := True
		}
	}else{
		try{
			word.Selection.ParagraphFormat.SpaceBeforeAuto := False
			word.Selection.ParagraphFormat.SpaceAfterAuto := false
		}
	}
	Word_Destroy()
	
}

/*
    函数:  Word_setParagraphs_Promote
    作用: 设置大纲级别-升级
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setParagraphs_Promote(){
	global word
	Word_Get()
	try
		word.Selection.Paragraphs.OutlinePromote
	Word_Destroy()
	
}

/*
    函数:  Word_setParagraphs_Demote
    作用: 设置大纲级别-降级
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_setParagraphs_Demote(){
	global word
	Word_Get()
	try
		word.Selection.Paragraphs.OutlineDemote
	Word_Destroy()
	
}

/*
    函数:  Word_PageSetup_Orientation
    作用: 设置页面方向
    参数: 
    返回:
    作者:  Kawvin
    版本:  0.1
	;枚举常量类型：WdOrientation
	wdOrientLandscape:=1
	wdOrientPortrait:=0
*/
Word_PageSetup_Orientation(){
	global word
	Word_Get()
	try
		word.Selection.PageSetup.Orientation:=!word.Selection.PageSetup.Orientation
	Word_Destroy()
	
}

/*
    函数:  Word_setMargin
    作用: 获取数据
    参数: fMargins
    返回:
    作者:  Kawvin
    版本:  0.1
	使用方法：Word_setMargin("1,1,1,1,0.5,0.5|2,2,2,2,1.5,1.5|3.17,3.17,2.54,2.54,1.5,1.75")
*/
Word_setMargin(fMargins){
	global MyDocLM,MyDocRM,MyDocTM,MyDocBM,MyDocHD,MyDocFD
	MyDocLM:=0,MyDocRM:=0,MyDocTM:=0,MyDocBM:=0,MyDocHD:=0,MyDocFD:=0
	GUI DocMargin:Destroy
	GUI DocMargin: -DPIScale
	GUI DocMargin:Default
	Gui DocMargin:Font, s12
	Gui DocMargin:Add, GroupBox, x12 y2 w542 h225, 页边距
	Gui DocMargin:Add, ListView, x22 y27 w521 h189 AltSubmit Grid gMyDocMarginsList, 上边距|下边距|左边距|右边距|页眉边距|页脚边距
	Gui DocMargin:Add, Text, x15 y235 w70 h29 +0x200, 上边距
	Gui DocMargin:Add, Text, x15 y275 w70 h29 +0x200, 下边距
	Gui DocMargin:Add, Text, x210 y235 w70 h29 +0x200, 左边距
	Gui DocMargin:Add, Text, x210 y275 w70 h29 +0x200, 右边距
	Gui DocMargin:Add, Text, x390 y235 w85 h29 +0x200, 页眉边距
	Gui DocMargin:Add, Text, x390 y275 w85 h29 +0x200, 页脚边距
	loop,Parse,fMargins,`|
	{
		if strlen(trim(A_LoopField))
			DocMArray:=StrSplit(A_LoopField,",")
		else
			continue
		if (A_index=1)
			MyDocTM:=DocMArray[1],MyDocBM:=DocMArray[2],MyDocLM:=DocMArray[3],MyDocRM:=DocMArray[4],MyDocHD:=DocMArray[5],MyDocFD:=DocMArray[6]
		LV_Add("",DocMArray[1],DocMArray[2],DocMArray[3],DocMArray[4],DocMArray[5],DocMArray[6])
	}
	LV_ModifyCol(1," center")
	LV_ModifyCol(2," center")
	LV_ModifyCol(3," center")
	LV_ModifyCol(4," center")
	LV_ModifyCol(5," center")
	LV_ModifyCol(6," center")
	Gui DocMargin:Add, Edit, x90 y233 w70 h31 center  vMyDocTM,%MyDocTM%
	Gui DocMargin:Add, Edit, x90 y273 w70 h31 center  vMyDocBM,%MyDocBM%
	Gui DocMargin:Add, Edit, x280 y233 w70 h31 center  vMyDocLM,%MyDocLM%
	Gui DocMargin:Add, Edit, x280 y273 w70 h31 center  vMyDocRM,%MyDocRM%
	Gui DocMargin:Add, Edit, x480 y233 w70 h31 center  vMyDocHD,%MyDocHD%
	Gui DocMargin:Add, Edit, x480 y273 w70 h31 center  vMyDocFD,%MyDocFD%
	Gui DocMargin:Add, Button, x120 y325 w100 h40 Default gMyDocMarginsOK, 确定
	Gui DocMargin:Add, Button, x350 y325 w100 h40 gMyDocMarginsCancel, 取消
	Gui DocMargin:Add, Text, x30 y375  h29 +0x200, 单击选中修改，双击直接设置
	Gui DocMargin:Font
	Gui DocMargin:Show, w570 h410, 页边距设置
	Return
	
	MyDocMarginsOK:
		GUI DocMargin:Default
		guicontrolget,MyDocTM,,MyDocTM
		guicontrolget,MyDocBM,,MyDocBM
		guicontrolget,MyDocLM,,MyDocLM
		guicontrolget,MyDocRM,,MyDocRM
		guicontrolget,MyDocHD,,MyDocHD
		guicontrolget,MyDocFD,,MyDocFD
		Word_PageSetup_Margin(MyDocTM,MyDocBM,MyDocLM,MyDocRM,MyDocHD,MyDocFD)
		GUI DocMargin:Destroy
	return
	
	MyDocMarginsCancel:
		GUI DocMargin:Destroy
	return
	
	MyDocMarginsList:
		GUI DocMargin:Default
		LV_GetText(MyDocTM, A_EventInfo,1) 
		LV_GetText(MyDocBM, A_EventInfo,2) 
		LV_GetText(MyDocLM, A_EventInfo,3) 
		LV_GetText(MyDocRM, A_EventInfo,4) 
		LV_GetText(MyDocHD, A_EventInfo,5) 
		LV_GetText(MyDocFD, A_EventInfo,6) 
		if A_GuiEvent = DoubleClick
		{
			Word_PageSetup_Margin(MyDocTM,MyDocBM,MyDocLM,MyDocRM,MyDocHD,MyDocFD)
			GUI DocMargin:Destroy
			Sleep 2000
			WinActivate ahk_class OpusApp
			Sleep 200
			SendLevel 1
			SendInput {Esc}
		}
		if A_GuiEvent = normal
		{
			guicontrol,,MyDocTM,%MyDocTM%
			guicontrol,,MyDocBM,%MyDocBM%
			guicontrol,,MyDocLM,%MyDocLM%
			guicontrol,,MyDocRM,%MyDocRM%
			guicontrol,,MyDocHD,%MyDocHD%
			guicontrol,,MyDocFD,%MyDocFD%
			Sleep 2000
			WinActivate ahk_class OpusApp
			Sleep 200
			SendLevel 1
			SendInput {Esc}
		}
	return
}

/*
    函数:  Word_PageSetup_Margin
    作用: 设置页面边距，页眉、页脚边距
    参数: fTM,fBM,fLM,fRM，上，下，左，右边距，fHD，设置页眉边距，fFD，设置页脚边距
    返回:
    作者:  Kawvin
    版本:  0.1
*/
Word_PageSetup_Margin(fTM:="",fBM:="",fLM:="",fRM:="",fHD:="",fFD:=""){
	global VimD
	global word
	Word_Get()
	;~ msgbox % CentimetersToPoints(fTM)
	;~ return
	if(fTM!="")
		;~ try
			word.Selection.PageSetup.TopMargin:=fTM*28.3527	;CentimetersToPoints(fTM)
	if(fBM!="")
		;~ try
			word.Selection.PageSetup.BottomMargin:=fBM*28.3527	;CentimetersToPoints(fBM)
	if(fLM!="")
		;~ try
			word.Selection.PageSetup.LeftMargin:=fLM*28.3527	;CentimetersToPoints(fLM)
	if(fRM!="")
		;~ try
			word.Selection.PageSetup.RightMargin:=fRM*28.3527	;CentimetersToPoints(fRM)
	if(fHD!="")
		;~ try
			word.Selection.PageSetup.HeaderDistance:=fHD*28.3527	;CentimetersToPoints(fHD)
	if(fFD!="")
		;~ try
			word.Selection.PageSetup.FooterDistance:=fFD*28.3527	;CentimetersToPoints(fFD)
	Word_Destroy()

}

/*
    函数:  CentimetersToPoints
    作用: 厘米转换为磅
    参数: fcm
    返回:
    作者:  Kawvin
    版本:  0.1
	1磅=1/72英寸=2.54/72厘米=0.03528厘米=0.3528毫米，1厘米=28.3527磅
*/
CentimetersToPoints(fcm){
	fPoints:=0
	if(fcm/1>0)
		fPoints:=fcm*28.3527
	return fPoints
}

/*
    函数:  Word_Insert_Date
    作用: 日期格式
    参数: 
    返回: 
    作者:  Kawvin
    版本:  0.1
*/
Word_Insert_Date(){
global word
	Word_Get() 
	word.Selection.InsertDateTime(DateTimeFormat:="yyyy'年'M'月'd'日'")
}


/*
    函数:  Word_settabstops
    作用: 设置制表位
    参数: Indent	，0.35=2字符
    返回:
    作者:  by 无关风月
    版本:  0.1
*/
Word_settabstops(){
	;~ if (Indent="")
		;~ return
	global word
	Word_Get()
	try
		{
		word.Selection.ParagraphFormat.tabstops.add(CentimetersToPoints(8.33),1,0)
		;~ word.Selection.ParagraphFormat.RightIndent := CentimetersToPoints(fFirstLineIndent)
	}
	Word_Destroy()
	
}
