<%@LANGUAGE=VBScript CodePage=65001%>
<%
Option Explicit
Response.Buffer=True
Response.addHeader "Content-Type","text/html; charset=utf-8"
Public K,SQL,FSO,XML,GlobalTime,GlobalRedirect,GlobalCounter
GlobalTime=Timer()
GlobalRedirect=False
GlobalCounter=0

'Title: FastSimple Kernel Class
'Author: Leo Amos
'Date: 2019/12/9
'所有系统常量以Global_Constant格式命名
'所有系统变量以GlobalVariable格式命名
'所有对象实例不赋予前缀，直接命名
Const Global_Version="3.2.5"

Class createKernel
	Private Map,App,XML,FSO,DB
	Private mySN,myPage,myRoot,myAppFolder,myAppRoot,myAppPage,myDB,myCookies
	Private myLoad,myLanguage,myTemplate,mySkin,myStyle,myTheme,myDefault,myMaster
	Private myAjax,myDisplay,myHTML,myMasterPack,myMasterCount,myPagePack,myPageCount
	Private myContent,myInside,myPreset(),myPresetSite

	Private Sub Class_Initialize()
		Install ""
	End Sub
	Private Sub Class_Terminate()
		Install "Clear"
	End Sub

	'内部处理函数和过程，用以解析和配置系统参数等
	Private Sub Install(IName)
		Select Case IName
		Case ""
			'设置系统标识码，保持唯一性可以保证稳定性和多系统并行
			If Global_Token="" Then
				mySN=Application("Global-Token")
				If mySN="" Then
					mySN=Random(8,"S")
					Application.Lock
					Application("Global-Token")=mySN
					Application.unLock
				End If
			Else
				mySN=uCase(Global_Token)
			End If

			'获取当前页面地址
			myPage=String(Request.serverVariables("PATH_INFO"))
			If inStrRev(myPage,".asp")=0 Then IName=Request.serverVariables("PATH_TRANSLATED"):myPage=myPage & Mid(IName,inStrRev(IName,"\")+1)

			'设置MAP对象
			If Not Global_Test And isBe(Application("Global-" & mySN)) Then
				Set Map=Application("Global-" & mySN)
				myRoot=Map.getAttribute("root")
			Else
				myRoot=myPage
				IName="/library/kernel.config"
				Set XML=Object("XML:.3.0")
				Do While inStr(myRoot,"/")>0
					myRoot=Mid(myRoot,1,inStrRev(myRoot,"/")-1)
					XML.Load Server.mapPath(myRoot & IName)
					If XML.parseError.errorCode=0 Then
						Set Map=XML.documentElement
						Install "Application"
						Exit Do
					Else
						If myRoot="" Then Install "1001"
					End If
				Loop
				Set XML=Nothing
			End If

			'检测APP对象并虚拟化页面路径
			setApp myRoot & getPath(Map,"app/")

			'设置系统参数
			myDB="access"
			myCookies=getConfig("cache.cookies")

			'设置HTML组件的必要参数
			myAjax=Request.queryString("random")
			If myAjax="" Then
				myAjax=False
				myDisplay=Request.queryString("display")
				If isNumeric(myDisplay) Then myDisplay=Fix(myDisplay) Else myDisplay=0
			Else
				myAjax=True
				myDisplay=-1
			End If
			myLoad=False
			myPageCount=-1
			myMasterCount=-1
			myInside=0
			myPresetSite=-1
			Redim myPreset(5,1)
		Case "Application"
			'简化缓存数据，删除说明文字和无效项目
			For Each IName In Map.selectNodes("//*[@intro]")
				IName.removeAttribute "intro"
			Next
			For Each IName In Map.selectNodes("*//map[not(@value)]")
				IName.setAttribute("value")=""
			Next
			For Each IName In Map.selectNodes("*//*[not(@name) || @name='']")
				IName.parentNode.removeChild IName
			Next
			'虚拟化MAP树
			Set FSO=Object("FSO")
			setRouter Map,""
			Set FSO=Nothing
			'写入根路径地址、补全结构
			Map.setAttribute("root")=myRoot
			Map.appendChild Map.parentNode.createElement("safe")
			Map.appendChild Map.parentNode.createElement("cache")
			'写入缓存对象
			If Global_Test Then Exit Sub
			Application.Lock
			Set Application("Kernel-" & mySN)=Map
			Application.unLock
		Case "Clear"
			If myLoad And Not GlobalRedirect Then Compiler:Response.Write myHTML
			If isObject(DB) Then
				DB.Close()
				Set DB=Nothing
			End If
			If isObject(App) Then Set App=Nothing
			Set Map=Nothing
		Case Else
			'请注意：虽已强制终止进程，但析构函数依然会被执行
			Response.Write "Found Error [" & IName & "]"
			Response.End
		End Select
	End Sub
	Private Sub setApp(IPath)
		If inStr(myPage,IPath)=1 Then
			'删除APP路径
			myPage=Mid(myPage,Len(IPath)+1)
			If inStr(myPage,"/")>0 Then
				'截取APP文件夹名称和内部页面路径
				myAppFolder=Mid(myPage,1,inStr(myPage,"/")-1)
				myAppPage=Mid(myPage,inStr(myPage,"/"))
				'取得APP真实根路径
				myAppRoot=IPath & myAppFolder
				'加载APP的开发者配置文件
				Set XML=Object("XML")
				XML.Load Server.mapPath(myAppRoot & "/library/developer.config")
				If XML.parseError.errorCode<>0 Then XML.loadXML "<dev/>"
				Set App=XML.documentElement
				If isVoid(App.getAttribute("title")) Then App.setAttribute("title")=myAppFolder
				Set XML=Nothing
			End If
			'补全APP路径
			myPage="/app/" & myPage
		Else
			'删除根路径
			myPage=Mid(myPage,Len(myRoot)+1)
		End If
		
	End Sub
	Private Sub setRouter(Parent,IPath)
		Dim Node,iName,IValue
		For Each Node In Parent.selectNodes("map")
			iName=String(Node.getAttribute("name"))
			iValue=String(Node.getAttribute("value"))
			Select Case Left(iValue,1)
			Case ""
				iValue=Parent.getAttribute("value") & "/" & iName
			Case "/"
				If iValue="/" Then iValue=""
			Case Else
				iValue=Parent.getAttribute("value") & "/" & iValue
			End Select
			Node.setAttribute("name")=iName
			Node.setAttribute("value")=iValue

			'对部分设置信息进行检测和更新
			iName=IPath & "/" & iName
			iValue=Server.mapPath(myRoot & iValue)
			If FSO Is Nothing Then
				'没有FSO权限仅加载设置文件
				If iName="/data/config" Then setConfig iValue
			Else
				'验证真实路径
				Select Case iName
				Case "/data"
					Node.setAttribute("value")="/" & getDataFolder("data")
				Case "/data/config"
					If FSO.folderExists(iValue) Then setConfig iValue
				Case "/res/user/language/default","/res/user/template/default","/res/user/skin/default"
					If Not FSO.folderExists(iValue) Then Node.setAttribute("value")="/res/" & Parent.getAttribute("name")
				Case Else
					If Not FSO.folderExists(iValue) Then FSO.createFolder iValue
				End Select 
			End If

			'如果存在子项则回调函数
			If Node.selectNodes("map").Length>0 Then setRouter Node,iName
		Next
	End Sub
	Private Function getDataFolder(IData)
		Dim Folder,iPath,iName

		If Left(IData,1)="$" Then
			IData=Mid(IData,2)
			iPath=myAppRoot
		Else
			If myRoot="" Then iPath="/" Else iPath=myRoot
		End If

		iPath=Server.mapPath(iPath)
		For Each Folder In FSO.getFolder(iPath).subFolders
			iName=String(Folder.Name)
			If inStr(iName,IData)>0 Then
				If iName<>IData Then getDataFolder=iName:Exit Function
				Exit For
			Else
				iName=""
			End If
		Next

		getDataFolder=getDataName(IData)
		Select Case iName
		Case ""
			FSO.createFolder iPath & "\" & getDataFolder
		Case IData
			FSO.moveFolder iPath & "\" & IData,iPath & "\" & getDataFolder
		End Select
	End Function
	Private Function getDataName(IName)
		Select Case Random(1,"+")
		Case 1,2,3
			getDataName=IName & "_" & Random(8,"")
		Case 4,5,6
			getDataName=Random(4,"") & "_" & IName & "_" & Random(4,"")
		Case 7,8,9
			getDataName=Random(8,"") & "_" & IName
		End Select
	End Function
	Private Sub setConfig(IPath)
		Dim XML,Self
		Set XML=Object("XML")
		For Each Self In Map.selectNodes("map[@name='data']/map[@name='config']/config")
			XML.Load IPath & "\" & Self.getAttribute("name") & ".xml"
			If XML.parseError.errorCode=0 Then setAttribute Self,XML.documentElement.Attributes
		Next
		Set Self=Nothing
		Set XML=Nothing
	End Sub
	Private Sub setAttribute(Self,Attributes)
		Dim Node,Item,iName
		For Each Item In Attributes
			iName=lCase(Item.Name)
			If iName<>"name" Then
				Select Case Self.getAttribute("name")
				Case "developer"
					'开发者允许添加更多的信息。
					Self.setAttribute(iName)=Item.Value
				Case Else
					Set Node=Self.selectSingleNode("@" & iName)
					If Not Node Is Nothing Then Node.Text=Item.Value:Set Node=Nothing
				End Select
			End If
		Next
		Set Item=Nothing
	End Sub
	Private Function getConfig(IName)
		Dim Node,iClass
		If inStr(IName,".")>1 Then
			iClass=Cut(IName,".")
			IName=Cut(IName,".\R")
		End If
		If iClass="" Then iClass="global"

		Set Node=Map.selectSingleNode("map[@name='data']/map[@name='config']/config[@name='" & iClass & "']/@" & IName)
		If Not Node Is Nothing Then getConfig=Node.Text:Set Node=Nothing
	End Function
	Private Function getPath(Parent,IPath)
		Dim Node,iPlace,iName,iValue
		getPath=Parent.getAttribute("value") & "/" & IPath
		If IPath="" Then Exit Function

		iPlace=inStr(IPath,"/")
		If iPlace>0 Then
			iName=Mid(IPath,1,iPlace-1)
		Else
			If inStr(IPath,".")>0 Or inStr(IPath,"?")>0 Or inStr(IPath,"#")>0 Then Exit Function Else iName=IPath
		End If

		Set Node=Parent.selectSingleNode("map[@name='" & lCase(iName) & "']")
		If Node Is Nothing Then Exit Function
		If iPlace>0 Then getPath=getPath(Node,Mid(IPath,iPlace+1)) Else getPath=Node.getAttribute("value")
		Set Node=Nothing
	End Function
	'供外部访问的超级方法
	Public Property Get Value(IName)
		If isVoid(IName) Then Exit Property Else IName=Trim(IName)

		Select Case Left(IName,1)
		Case "/"
			Value=myRoot & getPath(Map,Mid(IName,2))
		Case "^"
			'解析路径映射，但是不添加根路径。
			Value=Mid(IName,2)
			If Left(Value,1)="/" Then Value=getPath(Map,Mid(Value,2))
		Case "~"
			Value=Mid(IName,2)
			If Value="" Then Exit Property
			If Left(Value,1)="/" Then Value=Value(Value)
			Value=Server.mapPath(Value)
		Case Else
			Dim iFormat
			If inStr(IName,",")>1 Then
				iFormat=Cut(IName,",\RL")
				IName=Cut(IName,",\L")
			Else
				IName=lCase(IName)
			End If
			Value=getValue(IName)
			If iFormat<>"" Then Value=Format(Value,iFormat)
		End Select
	End Property
	Private Function getValue(IName)
		If inStr(IName,".")>1 Then
			getValue=Cut(IName,".")
			IName=Cut(IName,".\R")
			If IName="" Then Exit Function
			Select Case getValue
			Case "cookies"
				getValue=Cookies(IName)
			Case "session"
				getValue=Cache("session." & IName)
			Case "cache"
				getValue=Cache(IName)
			Case "form"
				getValue=Request.Form(IName)
				If myAjax Then getValue=Encode(getValue,4)
			Case "query"
				getValue=Request.queryString(IName)
				If myAjax Then getValue=Encode(getValue,4)
			Case "request"
				getValue=Request(IName)
				If myAjax Then getValue=Encode(getValue,4)
			Case "app"
				Select Case IName
				Case "folder"
					getValue=myAppFolder
				Case "page"
					getValue=myAppPage
				Case "root"
					getValue=myAppRoot
				Case Else
					If isObject(App) Then
						getValue=App.getAttribute(IName)
						If isVoid(getValue) Then getValue=""
					End If
				End Select
			Case "random"
				getValue=getRandom(IName)
			Case "config"
				getValue=getConfig(IName)
			Case "developer"
				getValue=getConfig("developer." & IName)
			Case "version"
				getValue=Version(IName)
			Case Else
				getValue=""
			End Select
		Else
			Select Case IName
			Case "root"
				getValue=myRoot
			Case "page"
				getValue=myPage
			Case "data","db"
				getValue=getFolder(IName)
			Case "from"
				getValue=Request.ServerVariables("HTTP_REFERER")
				If getValue="" Then
					If myRoot="" Then getValue="/" Else getValue=myRoot
				Else
					getValue=Replace(getValue,"%20"," ")
				End If
			Case "source"
				getValue=getSource()
			Case "go"
				getValue=String(Request("go"))
			Case "ajax"
				getValue=myAjax
			Case "sn"
				getValue=mySN
			Case "ip"
				getValue=IP()
			Case "version"
				getValue=Version("")
			End Select
		End If
	End Function
	Private Function getFolder(IName)
		If IName<>"data" Then getFolder="/" & IName
		If myAppFolder="" Then getFolder=Value("/data" & getFolder):Exit Function Else IName=myAppFolder
		Dim Node
		Set Node=Map.selectSingleNode("safe/*[@name='" & IName & "']/@value")
		If Node Is Nothing Then
			Set FSO=Object("FSO")
			If FSO Is Nothing Then
				IName="data"
			Else
				Set Node=Map.parentNode.createElement("s")
				Node.setAttribute("name")=IName
				IName=getDataFolder("$data")
				Node.setAttribute("value")=IName
				Map.selectSingleNode("safe").appendChild Node
				Set FSO=Nothing
			End If
		Else
			IName=Node.Text
		End If
		Set Node=Nothing
		getFolder=myAppRoot & "/"  & IName & getFolder
	End Function
	Private Function getSource()
		'来源地址被访问后立即删除
		getSource=Cache("session.source")
		If getSource="" Then getSource="from" Else Cache("session.source")=""
		'来源地址转换成真实路径后再反馈给用户
		getSource=Value(getSource)
		'某些来源地址不允许被重复访问，所以自动重置为首页
		If inStr(getSource,"/error/")>0 Or inStr(getSource,"/error.asp")>0 Or inStr(getSource,"/success/")>0 Or inStr(getSource,"/success.asp")>0 Or inStr(getSource,"/login/")>0 Or inStr(getSource,"/login.asp")>0 Or inStr(getSource,"/register/")>0 Or inStr(getSource,"/register.asp")>0 Then getSource=myRoot & "/"
	End Function
	Public Property Let Value(IName,Byval IValue)
		If isVoid(IName) Then Exit Property Else IName=Trim(IName)

		If inStr(IName,",")>1 Then
			IValue=Format(IValue,Cut(IName,",\RL"))
			IName=Cut(IName,",\L")
		Else
			IName=lCase(IName)
		End If

		If inStr(IName,".")>1 Then
			Value=Cut(IName,".")
			IName=Cut(IName,".\R")
			Select Case Value
			Case "cookies"
				Cookies(IName)=IValue
			Case "cache"
				Cache(IName)=IValue
			Case "session"
				Cache("session." & IName)=IValue
			End Select
		Else
			Select Case IName
			Case "redirect"
				GlobalRedirect=True
				If isVoid(IValue) Then
					IValue=getValue("from")
				Else
					IValue=Trim(IValue)
					If Left(IValue,1)="/" Then
						IValue=Value(IValue)
					Else
						If lCase(IValue)="source" Then IValue=getSource()
					End If
				End If
				Response.Redirect Replace(IValue,"/index.asp","/")
			Case "source"
				If isVoid(IValue) Then
					IValue=Request.ServerVariables("QUERY_STRING")
					If IValue="" Then IValue=myPage Else IValue=myPage & "?" & IValue
				Else
					If lCase(Trim(IValue))="remove" Then IValue=""
				End If
				Cache("session.source")=IValue
			Case "ajax"
				If Not myAjax Then
					'非Ajax请求的反馈设置
					Select Case String(IValue)
					Case "void" '不执行任何动作，停留在当前页
					Case Else
						'自动跳转到设定的页面，输入特殊值Source则自动跳转到来源页面
						setValue "redirect",IValue
					End Select
				Else
					'写入Ajax页面执行成功的值用以页面反馈
					Response.Write getConfig("ajax")
				End If
			End Select
		End If
	End Property
	Public Function Format(IValue,IType)
		If isVoid(IValue) Then Exit Function Else Format=Trim(IValue)
		If isVoid(IType) Then Exit Function Else IType=lCase(Trim(IType))

		If inStr(IType,",")>0 Then
			For Each IType In Split(IType,",")
				Format=Format(Format,IType)
			Next
		Else
			If isNumeric(IType) Then
				Format=Convert(Format,Fix(IType))
			Else
				If inStr(IType,".")>1 Then
					Dim iClass
					iClass=Cut(IType,".") '请勿删除iClass变量，部分输入变量名为iValue的值时会导致原值被修改
					IType=Cut(IType,".\R")
					Select Case iClass
					Case "encode"
						Format=Filter(Format,IType)
					Case "clip"
						Format=Clip(Format,IType)
					Case "interval"
						Format=Interval(Format,IType)
					Case "rule"
						Format=Rule(Format,IType)
					End Select
				Else
					Select Case String(IType)
					Case "length"
						Format=Length(Format)
					Case "char"
						Format=Length(Format)
					Case "l"
						Format=lCase(Format)
					Case "u"
						Format=uCase(Format)
					End Select
				End If
			End If
		End If
	End Function
	'设置、读取Cookies和Cache对象，IName必须使用小写
	Private Property Get Cookies(IName)
		Select Case inStr(IName,".")
		Case 0
			Cookies=mySN
		Case 1,Len(IName)
			Exit Property
		Case Else
			Cookies=Cut(IName,".")
			IName=Cut(IName,".\R")
			'自动根据APP环境设置分类名
			If Cookies="app" Then
				If myAppFolder="" Then Cookies="SERVICE" Else Cookies="APP-" & myAppFolder
			End If
			Cookies=mySN & "-" & Cookies
		End Select
		Select Case myCookies
		Case 1
			Cookies=Session(Cookies & "_" & IName)
		Case Else
			Cookies=Request.Cookies(Cookies)(IName)
		End Select
	End Property
	Private Property Let Cookies(IName,Byval IValue)
		Select Case inStr(IName,".")
		Case 0
			Cookies=mySN
		Case 1,Len(IName)
			Exit Property
		Case Else
			Cookies=Cut(IName,".")
			IName=Cut(IName,".\R")
			'自动根据APP环境设置分类名
			If Cookies="app" Then
				If myAppFolder="" Then Cookies="SERVICE" Else Cookies="APP-" & myAppFolder
			End If
			Cookies=mySN & "-" & Cookies
		End Select

		Select Case myCookies
		Case 1
			Select Case IName
			Case "$path","$expires","$life"
				Exit Property
			Case Else
				'非文本类的值将被删除
				IName=Cookies & "_" & IName
				If isVoid(IValue) Then Session.Contents.Remove IName Else Session(IName)=IValue
			End Select
		Case Else
			'禁止Cookies对象保存文本类的值
			If isVoid(IValue) Then IValue=""
			Select Case IName
			Case "$path"
				If IValue="" Then
					IValue=myRoot & "/"
				Else
					'如果输入的值不是绝对路径或者非路径则跳出设置
					If Left(IValue,1)="/" Then IValue=Value(IValue) Else Exit Property
					If Right(IValue,1)<>"/" Then IValue=IValue & "/"
				End If
				Response.Cookies(Cookies).Path=IValue
			Case "$expires"
				If isNumeric(IValue) Then
					IValue=Fix(IValue)+Now()
				Else
					If Not isDate(IValue) Then Exit Property
				End If
				Response.Cookies(Cookies).Expires=IValue
			Case "$life"
				IName=mySN & Replace(IName,"$","-")
				If IValue="" Then
					'输入值为空则只读取而不修改生命周期的值
					IValue=Request.Cookies(Cookies)(IName)
					If isDate(IValue) Then Response.Cookies(Cookies).Expires=IValue
				Else
					'如果是输入值为数字或日期格式则更新生命周期，否则将清空生命周期的值
					If isNumeric(IValue) Then IValue=Fix(IValue)+Now()
					If isDate(IValue) Then
						Response.Cookies(Cookies)(IName)=IValue
						Response.Cookies(Cookies).Expires=IValue
					Else
						Response.Cookies(Cookies)(IName)=""
					End If
				End If
				Response.Cookies(Cookies).Path=myRoot & "/"
			Case Else
				Response.Cookies(Cookies)(IName)=IValue
			End Select
		End Select
	End Property
	Private Property Get Cache(IName)
		Select Case inStr(IName,".")
		Case 0
		Case 1,Len(IName)
			Exit Property
		Case Else
			Cache=Cut(IName,".")
			IName=Cut(IName,".\R")
		End Select
		If Cache<>"memory" Then IName=mySN & "-" & IName

		Select Case Cache
		Case "session"
			If isObject(Session(IName)) Then Set Cache=Session(IName) Else Cache=Session(IName)
		Case "force"
			If isObject(Application(IName)) Then Set Cache=Application(IName) Else Cache=Application(IName)
		Case "life"
			Cache=Application(IName & "-time")
		Case "memory"
			Set Cache=Map.selectSingleNode("cache/*[@name='" & IName & "']")
		Case Else
			If isObject(Application(IName)) Then
				If getCacheLife(IName) Then Set Cache=Application(IName) Else Set Cache=Nothing
			Else
				If getCacheLife(IName) Then Cache=Application(IName) Else Cache=""
			End If
		End Select
	End Property
	Private Property Let Cache(IName,Byval IValue)
		Select Case inStr(IName,".")
		Case 0
		Case 1,Len(IName)
			Exit Property
		Case Else
			Cache=Cut(IName,".")
			IName=Cut(IName,".\R")
		End Select

		Select Case Cache
		Case "session"
			If IName="$expires" Then
				If isNumeric(IValue) Then
					IValue=Fix(IValue)
					If IValue>0 Then Session.timeOut=IValue
				End If
			Else
				IName=mySN & "-" & IName
				If isObject(IValue) Then
					Set Session(IName)=IValue
				Else
					If isArray(IValue) Then
						Session(IName)=IValue
					Else
						If isVoid(IValue) Then Session.Contents.Remove IName Else Session(IName)=IValue
					End If
				End If
			End If
		Case "life"
			IValue=setCacheLife(IValue)
			IName=mySN & "-" & IName & "-time"
			Application.Lock
			If isDate(IValue) Then Application(IName)=IValue Else Application.Contents.Remove IName
			Application.unLock
		Case "memory"
			Application.Lock
			Set Cache=Map.selectSingleNode("cache/*[@name='" & IName & "']")
			If Not Cache Is Nothing Then Cache.parentNode.removeChild Cache
			If isBe(IValue) Then
				On Error Resume Next
				Map.selectSingleNode("cache").appendChild IValue
				Err.Clear
			End If
			Application.unLock
		Case Else
			IName=mySN & "-" & IName
			Cache=setCacheLife(getConfig("cache.life"))
			Application.Lock
			If isObject(IValue) Then
				Set Application(IName)=IValue
			Else
				If isArray(IValue) Then
					Application(IName)=IValue
				Else
					IValue=Trim(IValue)
					If IValue="" Then
						Cache=""
						Application.Contents.Remove IName
					Else
						Dim iCommand
						iCommand=getConfig("command")
						If inStrRev(IValue,iCommand)>0 Then
							Cache=setCacheLife(Cut(IValue,iCommand & "\R"))
							IValue=Cut(IValue,iCommand)
						End If
						Application(IName)=IValue
					End If
				End If
			End If
			If isDate(Cache) Then Application(IName & "-time")=Cache Else Application.Contents.Remove IName & "-time"
			Application.unLock
		End Select
	End Property
	Private Function getCacheLife(IName)
		getCacheLife=Application(IName & "-time")
		If isDate(getCacheLife) Then
			If cDate(getCacheLife)>Now() Then getCacheLife=True Else getCacheLife=False
		Else
			getCacheLife=True
		End If
	End Function
	Private Function setCacheLife(IValue)
		If isDate(IValue) Then
			setCacheLife=cDate(IValue)
		Else
			If isVoid(IValue) Then Exit Function
			If isNumeric(IValue) Then
				IValue=Fix(IValue)
			Else
				If inStr(IValue,",")<>2 Then Exit Function
				setCacheLife=Mid(IValue,1,1)
				IValue=Format(Mid(IValue,3),1)
			End If
			'If IValue=0 Then Exit Function
			Select Case setCacheLife
			Case "s","n","h","d"
			Case Else
				setCacheLife="s"
			End Select
			setCacheLife=dateAdd(setCacheLife,IValue,Now)
		End If
	End Function
	'高级工具类
	Public Function Object(IValue)
		Select Case inStr(IValue,":")
		Case 0
			Object=uCase(Trim(IValue))
			IValue=""
		Case 1
			Set Object=Nothing
			Exit Function
		Case Else
			Object=Cut(IValue,"\U")
			IValue=Cut(IValue,"\R")
		End Select

		On Error Resume Next
		Select Case Object
		Case "XML"
			If IValue="" Then IValue=".3.0"
			Object="MSXML2.FreeThreadedDOMDocument" & IValue
		Case "SQL"
			Set Object=New createRecordSet
			If isBe(DB) Then
				'如果数据库已经连接成功，则从缓存中直接读取当前数据库。
				Object.Database("connect")=DB
				Object.Database("type")=myDB
			Else
				'连接默认数据库，连接成功后缓存该数据库。
				Object.Database("connect")=""
				Set DB=Object.Database("connect")
				myDB=Object.Database("type")
			End If
			If IValue<>"" Then Object.Run=IValue
		Case "SQL.APP"
			Set Object=New createRecordSet
			Object.Database("string")="$" & Global_DatabaseName
			If IValue<>"" Then Object.Run=IValue
		Case "FSO"
			If IValue="" Then Object="Scripting.FileSystemObject" Else Set Object=addFSO(IValue)
		Case "DEVELOPER"
			Set Object=Map.selectSingleNode("map[@name='data']/map[@name='config']/config[@name='developer']").Attributes
		Case "C"
			Object="ADODB.Connection"
		Case "RS"
			Object="ADODB.RecordSet"
		Case "S"
			Object="ADODB.Stream"
		End Select
		If isObject(Object) Then Exit Function
		Set Object=createObject(Object)
		If Err.Number<>0 Then Err.Clear:Set Object=Nothing
	End Function
	Private Function addFSO(IValue)
		Set addFSO=Object("FSO")
		If addFSO Is Nothing Then Exit Function

		Dim iType,iPath,iTarget,iKind
		If inStr(IValue,",")>1 Then
			IValue=Split(IValue,",")
			iType=uCase(IValue(0))
			If inStr(iType,".")>0 Then
				iKind=Cut(iType,".\R")
				iType=Cut(iType,".")
			End If
			Select Case uBound(IValue)
			Case 1
				iPath=IValue(1)
			Case 2
				iPath=IValue(1)
				iTarget=IValue(2)
			End Select
		Else
			iType="ADD"
			iPath=IValue
		End If
		iPath=Value("~" & iPath)
		If iKind="" Then
			If inStrRev(iPath,".")>0 Then iKind="FILE" Else iKind="FOLDER"
		End If

		Select Case iType
		Case "ADD"
			If iKind="FILE" Then
				If Not addFSO.fileExists(iPath) Then addFSO.CreateTextFile iPath
			Else
				If Not addFSO.folderExists(iPath) Then addFSO.createFolder iPath
			End If
		Case "REMOVE","DELETE"
			If iKind="FILE" Then
				If addFSO.fileExists(iPath) Then addFSO.deleteFile iPath
			Else
				If addFSO.folderExists(iPath) Then addFSO.deleteFolder iPath
			End If
		Case "COPY"
			iTarget=Value("~" & iTarget)
			If iKind="FILE" Then
				If addFSO.fileExists(iPath) Then addFSO.copyFile iPath,iTarget
			Else
				If addFSO.folderExists(iPath) Then addFSO.copyFolder iPath,iTarget
			End If
		Case "MOVE"
			iTarget=Value("~" & iTarget)
			If iKind="FILE" Then
				If addFSO.fileExists(iPath) And Not addFSO.fileExists(iTarget) Then addFSO.moveFile iPath,iTarget
			Else
				If addFSO.folderExists(iPath) And Not addFSO.folderExists(iTarget) Then addFSO.moveFolder iPath,iTarget
			End If
		End Select
	End Function
	'基础工具类，用以格式化数据和生成基础数据
	Private Function Convert(IValue,IType)
		Select Case IType
		Case 0 '转化为数值格式
			If isNumeric(IValue) Then Convert=IValue*1 Else Convert=0
		Case 1 '转化为整数格式1
			If isNumeric(IValue) Then Convert=Fix(IValue) Else Convert=0
		Case 2 '转化为整数格式2
			If isNumeric(IValue) Then Convert=Int(IValue) Else Convert=0
		Case 3 '转化为字节型数值0-255，小于0的重置为0，大于255的重置为255
			If isNumeric(IValue) Then
				Convert=Fix(IValue)
				If Convert<0 Then
					Convert=0
				Else
					If Convert>255 Then Convert=255
				End If
			Else
				Convert=0
			End If
		Case 4 '保留小数点2位
			If isNumeric(IValue) Then Convert=formatNumber(IValue,2,-1) Else Convert=0
		Case 5 '转化为布尔型的数值格式，-1和0，非零的数值全部格式化为-1
			If isNumeric(IValue) Then
				Convert=Fix(IValue)
				If Convert<>0 Then Convert=-1
			Else
				If Convert(IValue,6) Then Convert=-1 Else Convert=0
			End If
		Case 6 '转化为布尔值
			If isNumeric(IValue) Then
				If IValue=0 Then Convert=False Else Convert=True
			Else
				Convert=False
				If isVoid(IValue) Then Exit Function
				If lCase(Trim(IValue))="true" Then Convert=True
			End If
		Case 7 '过滤非日期格式
			If isDate(IValue) Then Convert=IValue
		Case 8 '重置非日期格式为当前时间
			If isDate(IValue) Then Convert=IValue Else Convert=Now()
		Case 9 '重置非日期格式为当前日期
			If isDate(IValue) Then Convert=dateValue(IValue) Else Convert=Date()
		Case 10 '格式化日期格式为标准模式 mm-dd-yyyy[ hh:mmmm:ss]
			If isDate(IValue) Then Convert=Month(IValue) & "-" & Day(IValue) & "-" & Year(IValue) & " " & timeValue(IValue)
			'Convert=Replace(Convert," 0:00:00","")
		Case 11 '格式化日期格式为标准日期模式 mm-dd-yyyy
			If isDate(IValue) Then Convert=Month(IValue) & "-" & Day(IValue) & "-" & Year(IValue)
		Case 20,50,100,255,10000 '截取指定长度的字符串
			Convert=Clip(IValue,IType)
		Case Else
			Convert=Trim(IValue)
		End Select
	End Function
	Private Function Encode(IValue,IType)
		If IValue="" Then Exit Function Else Encode=IValue

		Select Case lCase(IType)
		Case "-2" '
			Encode=Replace(Encode,"<br/>",Chr(10))
			Encode=Replace(Encode,"&nbsp;&nbsp;&nbsp;&nbsp;",Chr(9))
			Encode=Encode(IValue,-1)
		Case "-1","html" '恢复HTML标签、单双引号
			Encode=Replace(Encode,"&acute;",Chr(39))
			Encode=Replace(Encode,"&quot",Chr(34))
			Encode=Replace(Encode,"&lt;","<")
			Encode=Replace(Encode,"&gt;",">")
		Case "1","text" '转换HTML标签、单双引号
			Encode=Replace(Encode,">","&gt;")
			Encode=Replace(Encode,"<","&lt;")
			Encode=Replace(Encode,Chr(34),"&quot")
			Encode=Replace(Encode,Chr(39),"&acute;")
		Case "2","content" '保留制表符、和回车换行效果
			Encode=Encode(IValue,1)
			Encode=Replace(Encode,Chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
			Encode=Replace(Encode,Chr(10),"<br/>")
		Case "3","value" '纯文本模式，清空制表符、换行键、回车键等，禁止排版
			Encode=Encode(IValue,1)
			Encode=Replace(Encode,Chr(8),"") '回格
			Encode=Replace(Encode,Chr(9),"") 'Tab水平制表符
			Encode=Replace(Encode,Chr(10),"") '回车
			Encode=Replace(Encode,Chr(11),"") 'Tab垂直制表符
			Encode=Replace(Encode,Chr(12),"") '换页
			Encode=Replace(Encode,Chr(13),"") '换行
		Case "4","ajax" '转换AJAX请求的特殊字符
			Encode=Replace(Encode,"%3D","=")
			Encode=Replace(Encode,"%26","&")
			Encode=Replace(Encode,"%25","%")
		Case "5","url" '转换URL地址中的空格
			Encode=Replace(Encode,"%20"," ")
		End Select
	End Function
	Private Function Interval(IValue,IType)
		If isDate(IValue) Then
			Select Case IType
			Case "s","n","h","d","w","ww","m","q","y","yyyy"
				Interval=dateDiff(IType,IValue,Now())*1
			Case Else
				Interval=0
			End Select
		Else
			Interval=0
		End If
	End Function
	Private Function Rule(IValue,IType)
		Dim XML,Node
		Set XML=Object("XML")
		XML.Load Value("~/data/config/regexp.xml")
		If XML.parseError.errorCode=0 Then
			Set Node=XML.documentElement.selectSingleNode("*[@name='" & String(IType) & "']")
			If Not(Node Is Nothing) Then Rule=Node.Text
			Set Node=Nothing
		Else
			XML.loadXML("<regexp/>")
			XML.Save Value("~/data/config/regexp.xml")
		End If
		Set XML=Nothing

		'执行正则表达式的验证任务，表达式为空默认通过验证。
		If Rule="" Then
			Rule=True
		Else
			Dim Re
			Set RE=New RegExp
			RE.Pattern=Rule
			If RE.Test(IValue) Then Rule=True Else Rule=False
			Set RE=Nothing
		End If
	End Function
	Private Function getRandom(IValue)
		If isNumeric(IValue) Then
			getRandom=Random(IValue,"")
		Else
			If inStr(IValue,"\")>1 Then
				getRandom=Cut(IValue,"\\R")
				IValue=Cut(IValue,"\")
				getRandom=Random(IValue,getRandom)
			Else
				getRandom=Random(8,IValue)
			End If
		End If
	End Function
	Private Function Random(ILength,IType)
		Randomize
		Select Case uCase(IType)
		Case "L","1" '小写字母
			For IType=1 To ILength
				Random=Random & Chr(Cint(Rnd*25)+97)
			Next
		Case "U","2" '大写字母
			Random=uCase(Random(ILength,1))
		Case "E","3" '大小写字母
			For IType=1 To ILength
				Random=Random & Random(1,Cint(Rnd*1)+1)
			Next
		Case "+","4" '正数Positive，不包括零
			Random=Cint(Rnd*8)+1
			For IType=2 To ILength
				Random=Random & Cint(Rnd*9)
			Next
			Random=1*Random
		Case "-","5" '负数Minus，不包括零
			Random=-1*Random(ILength,4)
		Case "N","6" '自然数，包括零
			If ILength=1 Then Random=Cint(Rnd*9)*1 Else Random=Random(ILength,4)
		Case "I","7" '整数，包括正数、负数和零
			If Cint(Rnd*1)=1 Then IType=1 Else IType=-1
			Random=IType*Random(ILength,6)
		Case Else
			'随机数，不包括零以防止字母o和零之间产生认知错误
			Random=Random(1,3)
			For IType=2 To ILength
				Random=Random & Random(1,Cint(Rnd*1)+3)
			Next
		End Select
	End Function
	Private Function Char(IValue)
		Char=Abs(AscW(IValue))
		'If Char<0 Then Char=Char+65536
		If Char>255 Then Char=1 Else Char=0
	End Function
	Private Function Length(IValue)
		If IValue="" Then
			Length=0
		Else
			Length=Len(IValue)
			If Len("中文")=2 Then
				Dim i
				For i=1 To Length Step 1
					Length=Length+Char(Mid(IValue,i,1))
				Next
			End If
		End If
	End Function
	Private Function Clip(IValue,ILength)
		If IValue="" Then Exit Function
		ILength=Fix(ILength)
		If ILength<1 Then Exit Function
		If Len("中文")=2 Then
			Dim i,ii
			ii=Len(IValue)
			If ILength<ii*2 Then
				For i=1 To ii Step 1
					ii=Mid(IValue,i,1)
					ILength=ILength-Char(ii)-1
					If ILength<0 Then Exit For Else Clip=Clip & ii
				Next
			Else
				Clip=IValue
			End If
		Else
			Clip=Left(IValue,ILength)
		End If
	End Function
	Private Function IP()
		Dim iAgent
		iAgent=Request.ServerVariables("REMOTE_ADDR")
		IP=lCase(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))

		If IP="" Or inStr(IP,"unknown")>0 Then
			IP=iAgent
			iAgent=""
		elseIf inStr(IP,",")>0 Then
			IP=Mid(IP,1,inStr(IP,",")-1)
		elseIf inStr(IP,";")>0 Then
			IP=Mid(IP,1,inStr(IP,";")-1)
		End If

		IP=Trim(Mid(IP,1,30))
		If IP<>"" Then
			IP=Replace(IP,Chr(0),"")
			IP=Replace(IP,"'","''")
		End If
		If iAgent<>"" Then IP=IP & "/" & iAgent
	End Function
	Private Property Get Version(IName)
		Select Case IName
		Case ""
			Version=Global_Version & " Standard"
		Case "title"
			Version=getConfig("developer.core")
		Case "full"
			Version=Version("title") & " Version " & Version("")
		Case "author"
			Version="Leo Amos"
		Case "mail"
			Version="master@iamos.cn"
		Case "update"
			Version="http://fastsimple.cn/asp/update/?version=" & Version("")
		End Select
	End Property
	'HTML组件
	Public Property Get HTML(IName)
		If isVoid(IName) Then Exit Property Else IName=lCase(Trim(IName))
		If isNumeric(IName) Then
			HTML=getTemplate(IName)
		Else
			If inStr(IName,":")>1 Then
				HTML=Cut(IName,":\R")
				If HTML="" Then Exit Property
				IName=Cut(IName,":")

				Select Case IName
				Case "read"
					HTML=getFile(HTML)
				Case "page"
					If isNumeric(HTML) Then HTML=getTemplate(HTML)
				Case "main"
					If isNumeric(HTML) Then HTML=getTemplate(HTML)
				Case "cookies"
					If mySkin="" And myStyle="" Then
						mySkin=Cookies(HTML & ".skin")
						myStyle=Cookies(HTML & ".style")
					End If
					If myLanguage="" Then myLanguage=Cookies(HTML & ".language")
				End Select
			Else
				Select Case IName
				Case "add"
					loadHTML()
				Case "update"
					'更新内置模板对象，该属性需要配合Change方法共同使用
					If myInside=0 Then
						Exit Property
					elseIf myInside>0 Then
						If myInside<=myPageCount Then myPagePack(myInside)=myContent
					elseIf myInside<0 Then
						myInside=-1*myInside
						If myInside<=myMasterCount Then myMasterPack(myInside)=myContent
					End If
					myInside=0
					myContent=""
				Case "display"
					HTML=myDisplay
				Case "cookies"
					HTML IName & ":app"
				End Select
			End If
		End If
	End Property
	Public Property Let HTML(IName,Byval IValue)
		If isVoid(IName) Then Exit Property Else IName=lCase(Trim(IName))
		Select Case inStr(IName,".")
		Case 0
			Select Case IName
			Case "add"
				loadHTML()
				HTML("page")=IValue
			Case "page"
				If isNumeric(IValue) Then
					IValue=getTemplate(IValue)
				Else
					If isVoid(IValue) Then IValue=""
				End If
				myHTML=Replace(myHTML,"{{page}}",IValue)
			Case "display"
				If isNumeric(IValue) Then IValue=Fix(IValue) Else Exit Property
				Select Case IValue
				Case -1,1,2,3
					myDisplay=IValue
				Case Else
					myDisplay=0
				End Select
			Case "skin"
				IValue=String(IValue)
				If IValue="" Then Exit Property
				If inStr(IValue,",")>1 Then
					mySkin=Cut(IValue,",")
					myStyle=Cut(IValue,",\R")
				Else
					mySkin=IValue
					myStyle=""
				End If
			Case "style"
				myStyle=String(IValue)
			Case "language"
				myLanguage=String(IValue)
			End Select
		Case 1
			myHTML=Replace(myHTML,"{#" & Mid(IName,2) & "}",getCode(IValue))
		Case Else
			HTML=Cut(IName,".\R")
			If HTML="" Then Exit Property
			IName=Cut(IName,".")
			If isNumeric(IName) Then
				IName=Fix(IName)
				If IName=0 Then Exit Property
				'当更新对象为数字时则自动转换为内置模板对象
				If IName<>myInside Then
					HTML "update" '应用对象不同时自动更新
					myInside=IName
					myContent=HTML(myInside)
				End If
				If myContent="" Then Exit Property
				myContent=Replace(myContent,"{#" & HTML & "}",getCode(IValue))
			Else
				Select Case IName
				Case "code"
					myHTML=Replace(myHTML,"{#" & HTML & "}",getCode(IValue))
				Case "preset"
					If isVoid(IValue) Then IValue=""
					myPresetSite=myPresetSite+1
					If uBound(myPreset)<myPresetSite Then Redim Preserve myPreset(myPresetSite+5,1)
					myPreset(myPresetSite,0)=HTML
					myPreset(myPresetSite,1)=IValue
				Case "cookies"
					
				End Select
			End If
		End Select
	End Property
	Public Function Code(IContent,IName,IValue)
		If isNumeric(IContent) Then
			HTML(IContent & "." & IName)=IValue
		Else
			IValue=getCode(IValue)
			IName="{#" & String(IName) & "}"
			If isVoid(IContent) Then
				myHTML=Replace(myHTML,IName,IValue)
			Else
				Code=Replace(IContent,IName,IValue)
			End If
		End If
	End Function
	Private Function getTemplate(IValue)
		IValue=Fix(IValue)
		If IValue>0 Then
			If IValue<=myPageCount Then getTemplate=myPagePack(IValue)
		elseIf IValue<0 Then
			IValue=-1*IValue
			If IValue<=myMasterCount Then getTemplate=myMasterPack(IValue)
		elseIf IValue=0 Then
			getTemplate=myHTML
		End If
	End Function
	Private Function getCode(IValue)
		If isVoid(IValue) Then Exit Function Else IValue=Trim(IValue)
		Select Case Left(IValue,1)
		Case "^"
			If isNumeric(Mid(IValue,2)) Then IValue=getTemplate(Mid(IValue,2))
		Case "$"
			If isNumeric(Mid(IValue,2)) Then IValue="{$page(" & Mid(IValue,2) & ")}"
		End Select
		getCode=IValue
	End Function
	Private Function getFile(IName)
		Dim XML,iPage,iFolder
		If inStrRev(IName,".")=0 Then IName=IName & ".lng"
		If inStr(IName,",")>1 Then
			iFolder=Cut(IName,",")
			IName=Cut(IName,",\R")
			iPage=Value("~/res/user/language/" & myLanguage & "/app/" & iFolder & "/" & IName)
		Else
			iPage=Value("~/res/user/language/" & myLanguage & "/" & IName)
		End If

		Set XML=Object("XML")
		XML.Load iPage
		If XML.parseError.errorCode=0 Then
			getFile=XML.documentElement.Text
		Else
			If iFolder="" Then Exit Function
			XML.Load Value("~/app/" & iFolder & "/res/language/" & myLanguage & "/" & IName)
			If XML.parseError.errorCode=0 Then getFile=XML.documentElement.Text
		End If
		Set XML=Nothing
	End Function
	Private Sub loadHTML()
		'如果HTML组件已经加载或者已启动跳转任务则终止进程
		If myLoad Or GlobalRedirect Then Exit Sub

		Set XML=Object("XML")
		loadDefault()
		loadTemplate myTemplate,Replace(myPage,".asp",".xml")
		Select Case myDisplay
		Case -1 '不加载母版页
		Case 0 '分页模板加载失败或者不存在
			If Value("config.display.blank")=1 Then loadTemplate myTemplate,"/blank.master"
		Case 2 '强制加载系统母版页
			 myDisplay=1
		Case 1,3 '1标准模式，3只加载APP母版页
			If myAppFolder="" Then
				myDisplay=1
			Else
				loadTemplate myTemplate,"/app/" & myAppFolder & myMaster
				If Not myLoad And myDisplay=3 Then myDisplay=1
			End If
		End Select
		If myDisplay=1 Then loadTemplate myTemplate,myMaster
		Set XML=Nothing

		If myHTML<>"" Then
			If myPageCount>-1 Then myLoad=1 Else myLoad=0
			If myMasterCount>-1 Then myLoad=myLoad+2
			Select Case myLoad
			Case 1
				myPagePack=Split(myHTML,"$/$")
				myHTML=myPagePack(0)
			Case 2
				myMasterPack=Split(myHTML,"$/$")
				myHTML=Replace(myMasterPack(0),"{{0}}","")
			Case 3
				myLoad=Split(myHTML,"$\$")
				myPagePack=Split(myLoad(0),"$/$")
				myMasterPack=Split(myLoad(1),"$/$")
				myHTML=Replace(myMasterPack(0),"{{0}}",myPagePack(0))
			End Select
			If myPageCount>=0 Then myPagePack(0)=""
			If myMasterCount>=0 Then myMasterPack(0)=""
		End If

		myLoad=True
	End Sub
	Private Sub loadDefault()
		myDefault="default"
		myTemplate=myDefault
		If mySkin="" And myStyle="" Then
			mySkin=getConfig("display.skin")
			myStyle=getConfig("display.style")
		End If
		If mySkin="" Then mySkin=myDefault
		mySkin=Value("/res/user/skin/" & mySkin)
		XML.Load Server.mapPath(mySkin & "/setting.xml")
		If XML.parseError.errorCode=0 Then
			Dim Item
			For Each Item In XML.documentElement.Attributes
				Select Case Item.Name
				Case "template"
					myTemplate=String(Item.Value)
					If myTemplate="" Then myTemplate=myDefault
				Case "style"
					If myStyle="" Then myStyle=String(Item.Value)
				End Select
			Next
			Set Item=Nothing
		End If
		If myStyle="" Or myStyle=myDefault Then myStyle=mySkin
		If myLanguage="" Then
			myLanguage=getConfig("display.language")
			If myLanguage="" Then myLanguage=myDefault
		End If
		myMaster="/main.master"
	End Sub
	Private Sub loadTemplate(ITemplate,IPage)
		myLoad=False
		XML.Load Value("~/res/user/template/" & ITemplate  & IPage)
		If XML.parseError.errorCode=0 Then
			myLoad=True
		Else
			If ITemplate=myDefault Then
				If myAppPage="" Or IPage=myMaster Then Exit Sub
				If inStr(IPage,myMaster)=0 Then IPage=Replace(myAppPage,".asp",".xml") Else IPage=myMaster
				XML.Load Server.mapPath(myAppRoot & "/res/template" & IPage)
				If XML.parseError.errorCode=0 Then myLoad=True
			Else
				loadTemplate myDefault,IPage
			End If
		End If

		If Not myLoad Then Exit Sub

		If inStr(IPage,myMaster)>0 Then
			myMasterPack=XML.documentElement.Text
			If myMasterPack="" Then Exit Sub
			If myHTML="" Then myHTML=myMasterPack Else myHTML=myHTML & "$\$" & myMasterPack
			myMasterPack=Split(myMasterPack,"$/$")
			myMasterCount=uBound(myMasterPack) 
		Else
			If myDisplay=0 Then
				myDisplay=XML.documentElement.getAttribute("display")
				If Not isNumeric(myDisplay) Then myDisplay=1
			End If
			myPagePack=XML.documentElement.Text
			If myPagePack="" Then Exit Sub
			Dim Item
			For Each Item In XML.documentElement.Attributes
				myPagePack=Replace(myPagePack,"{{" & Item.Name & "}}",Item.Value)
			Next
			myHTML=myPagePack
			myPagePack=Split(myPagePack,"$/$")
			myPageCount=uBound(myPagePack)
		End If
	End Sub
	Private Sub Compiler()
		Dim i,Item

		'置换内置模板对象，如果内置模板有过修改记录则先更新内置模板
		If myInside<>0 Then HTML "update"
		For i=1 To myPageCount Step 1
			myHTML=Replace(myHTML,"{{" & i & "}}",myPagePack(i))
		Next
		For i=1 To myMasterCount Step 1
			myHTML=Replace(myHTML,"{[" & i & "]}",myMasterPack(i))
		Next
		'置换预置变量并且清空存储器
		For i=myPresetSite To 0 Step -1
			myHTML=Replace(myHTML,"{#" & myPreset(i,0) & "}",myPreset(i,1))
		Next
		Erase myPreset
		'置换开发者信息
		For Each Item In Object("developer")
			myHTML=Replace(myHTML,"{$dev." & Item.Name & "}",Item.Value)
		Next

		Set XML=Object("XML")
		'加载语言包
		loadLanguage()
		'加载主题设置
		loadTheme Server.mapPath(mySkin),Server.mapPath(myStyle)
		Set XML=Nothing

		myHTML=Replace(myHTML,"{$app.path}",myAppRoot)
		myHTML=Replace(myHTML,"{$app.title}",Value("App.Title"))
		myHTML=Replace(myHTML,"{$app}",Value("/app"))
		myHTML=Replace(myHTML,"{$style}",myStyle)
		myHTML=Replace(myHTML,"{$skin}",mySkin)
		myHTML=Replace(myHTML,"{$image}","{$theme}/image")
		myHTML=Replace(myHTML,"{$theme}",Value("/res/theme"))
		myHTML=Replace(myHTML,"{$service}","{$root}/service")
		myHTML=Replace(myHTML,"{$root}",Value("root"))
		myHTML=Replace(myHTML,"{$counter}",GlobalCounter)
		myHTML=Replace(myHTML,"{$runtime}",FormatNumber((Timer()-GlobalTime),5)*1000)
	End Sub
	Private Sub loadLanguage()
		Dim Lng
		Set Lng=Object("XML")
		Lng.Load Value("~/res/user/language/" & myLanguage & myMaster)
		If Lng.parseError.errorCode=0 Then
			'加载当前页面的语言包
			XML.Load Value("~/res/user/language/" & myLanguage & Replace(myPage,".asp",".xml"))
			If XML.parseError.errorCode=0 Then
				codeLanguage XML,"page"
				If myAppFolder<>"" Then
					XML.Load Value("~/res/user/language/" & myLanguage & "/app/" & myAppFolder & myMaster)
					If XML.parseError.errorCode=0 Then codeLanguage XML,"app"
				End If
			Else
				If myAppFolder<>"" Then
					XML.Load Server.mapPath(myAppRoot & "/res/language/" & myLanguage & Replace(myAppPage,".asp",".xml"))
					If XML.parseError.errorCode=0 Then codeLanguage XML,"page"
					XML.Load Server.mapPath(myAppRoot & "/res/language/" & myLanguage & myMaster)
					If XML.parseError.errorCode=0 Then codeLanguage XML,"app"
				End If
			End If
			'转换全局语言包
			codeLanguage Lng,"web"
		Else
			'如果主包不存在则自动切换回默认语言包
			If myLanguage<>myDefault Then
				myLanguage=myDefault
				loadLanguage
			End If
		End If
		Set Lng=Nothing
	End Sub
	Private Sub codeLanguage(XML,ITag)
		Dim Item,iData
		iData=XML.documentElement.Text
		If iData<>"" Then
			iData=Split(iData,Chr(10))
			For i=uBound(iData)+1 To 1 Step -1
				myHTML=Replace(myHTML,"{$" & ITag & "(" & i & ")}",iData(i-1))
			Next
		End If
		For Each Item In XML.documentElement.Attributes
			myHTML=Replace(myHTML,"{$" & ITag & "." & String(Item.Name) & "}",Item.Value)
		Next
		Set Item=Nothing
	End Sub
	Private Sub loadTheme(ISkin,IStyle)
		Dim Node,iUpdate
		XML.Load IStyle & "\theme.xml"
		If XML.parseError.errorCode=0 Then
			iUpdate=Value("config.theme.update,1")
			If iUpdate<0 Then iUpdate=24
			'默认设置Theme数据24小时更新一次
			If Interval(XML.documentElement.getAttribute("update"),"h")<24 Then
				For Each Node In XML.documentElement.childNodes
					myHTML=Replace(myHTML,"{$theme." & Node.getAttribute("name") & "}",Node.getAttribute("value"))
				Next
				Exit Sub
			End If
		End If

		'自动生成Theme数据，需要FSO权限支持，Theme和Style目录必须存在
		Set FSO=Object("FSO")
		If FSO Is Nothing Then Exit Sub
		If Not FSO.folderExists(IStyle) Then Exit Sub
		Dim Folder,iPath,iName
		iPath=Value("~/res/theme")
		If Not FSO.folderExists(iPath) Then Exit Sub
		'验证通过，开始生成数据
		XML.loadXML "<theme/>"
		For Each Folder In FSO.getFolder(iPath).subFolders
			iName=String(Folder.Name)
			Set Node=XML.createElement("item")
			Node.setAttribute("name")=iName
			Node.setAttribute("value")="{$theme}/" & iName
			XML.documentElement.appendChild Node
			Set Node=Nothing
		Next
		updateTheme ISkin,"skin"
		If IStyle<>ISkin Then updateTheme IStyle,"style"
		Set FSO=Nothing

		'保存数据更新时间并重载进程
		XML.documentElement.setAttribute("update")=Now()
		XML.Save IStyle & "\theme.xml"
		loadTheme ISkin,IStyle
	End Sub
	Private Sub updateTheme(IPath,ITag)
		'检测Skin或者Style路径下的Theme目录
		If Not FSO.folderExists(IPath & "\theme") Then Exit Sub
		Dim Folder,Node,iName
		For Each Folder In FSO.getFolder(IPath & "\theme").subFolders
			iName=String(Folder.Name)
			Set Node=XML.documentElement.selectSingleNode("*[@name='" & iName & "']")
			If Not Node Is Nothing Then
				Node.setAttribute("value")="{$" & ITag & "}/theme/" & iName
				Set Node=Nothing
			End If
		Next
	End Sub
End Class
'快速sql组件
Class createRecordSet
	Private RS,db,dbLock,dbType,dbString
	Private myCounter,myMode,mySafe,myTag,mySize,myCommand
	Private myAllRecord,myNowRecord,myAllPage,myNowPage,myPageTag,myPageSize
	Private Sub Class_Initialize()
		myCounter=0
		myMode=0
		mySafe=False
		myTag=K.Value("config.db.tag")
		mySize=K.Value("config.db.size")
		myCommand=K.Value("config.command")
	End Sub
	Private Sub Class_Terminate()
		Clear()
		'Database "close"
		GlobalCounter=GlobalCounter+myCounter
	End Sub
	Public Function Connect(IValue)
		dbLock=True '打开数据库类型锁
		'解析输入的字符串
		If Not isArray(IValue) Then
			If isVoid(IValue) Then IValue=K.Value("config.db.string")
			IValue=Split(IValue,",")
		End If
		If uBound(IValue)>2 Then dbType=1 Else dbType=0
		'生成数据库连接字符串
		Dim iString
		Select Case dbType
		Case 1
			iString="Provider=Sqloledb; Data Source=" & IValue(0) & "; Initial Catalog=" & IValue(1) & "; User ID=" & IValue(2) & "; Password=" & IValue(3) & ";"
		Case Else
			'个人ACCESS数据库说明：/a.mdb指向一个绝对路径；a.mdb指向系统数据库目录；b/a.mdb指向相对路径，$a.mdb指向APP数据库目录，如果不是APP应用则指向系统数据库目录。
			iString=Trim(IValue(0))
			Select Case Left(iString,1)
			Case "/"
				iString=K.Value(iString)
			Case "$"
				iString=K.Value("db") & "/" & Mid(iString,2)
			Case Else
				If inStr(iString,"/")=0 Then iString=K.Value("/data/db/" & iString)
			End Select
			iString=Server.mapPath(iString)

			If inStr(iString,".accdb")>0 Then
				iString="Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & iString & "; Persist Security Info=False;"
			Else
				iString="Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & iString & ";"
			End If
			If uBound(IValue)>0 Then iString=iString & " Jet OLEDB:Database Password=" & IValue(1) & ";"
		End Select

		'创建数据库连接对象，如果发生错误则返回错误信息
		On Error Resume Next
		Set Connect=K.Object("ADODB.Connection")
		Connect.Open iString
		If Err.Number=0 Then Exit Function
		iString=Err.Description & iString
		Err.Clear
		Set Connect=Nothing
		If Not BASETEST And inStr(iString,"\")>0 Then
			'对数据库路径进行过滤以防止数据目录泄露
			'iString=Mid(iString,1,inStr(iString,":")-2) & Mid(iString,inStrRev(iString,"\")+1)
			iString=Mid(iString,inStrRev(iString,"\")+1)
			If inStr(iString,";") Then iString=Mid(iString,1,inStr(iString,";")-1)
		End If
		Response.Write "Database Connect Is Error : " & iString
		Response.End
	End Function

	Public Property Let Database(IName,Byval IValue)
		Select Case String(IName)
		Case "string" '预设数据库连接的默认参数
			dbString=IValue
		Case "connect" '输入外部的数据库连接对象，传入成功后可以主动更改一次数据库类型。
			If isBe(IValue) Then
				Set db=IValue
				dbType=0     '默认设置数据库类型为ACCESS
				dbLock=False '关闭数据库类型锁
			Else
				Set db=Connect(IValue)
			End If
		Case "type"
			If Not dbLock And String(IValue)="mssql" Then myType=1
			myLock=True '打开数据库类型锁
		End Select
	End Property
	Public Property Get Database(IName)
		Select Case String(IName)
		Case "string" '使用预设参数连接数据库,如果预设参数不存在则使用系统参数，如果数据库已连接则忽略该方法
			If Not isBe(db) Then Set db=Connect(dbString)
		Case "type"
			If dbType=1 Then Database="mssql" Else Database="access"
		Case "connect" '输出当前已连接的数据库对象
			If isBe(db) Then Set Database=db
		Case "close"
			Clear()
			If isBe(db) Then
				db.Close
				Set db=Nothing
			End If
		End Select
	End Property

	'针对不同的数据库类型修正内置now()函数的使用，关键字“$now$”。
	Private Function Revision(IValue)
		IValue=Trim(IValue)
		Select Case dbType
		Case 1
			Revision=Replace(IValue,"$now$","GetDate()")
		Case Else
			Revision=Replace(IValue,"$now$","Now()")
		End Select
	End Function
	Public Property Let Execute(Byval IString)
		On Error Resume Next
		db.Execute Revision(IString)
		If Err.Number=0 Then Exit Property
		If Global_Test Then Response.Write IString & "<br/>"
		Response.Write "SQL Error:" & Err.Source & "[" & Err.Number & "]；" & Err.Description
		Err.Clear
	End Property
	Public Property Let Run(Byval IString)
		On Error Resume Next
		Database "String"
		IString=Revision(IString)
		Select Case uCase(Left(IString,6))
		Case "SELECT"
			Clear() '清理上一次查询对象
			If inStrRev(IString,myCommand)>0 Then
				Run=Mid(IString,inStrRev(IString,myCommand)+Len(myCommand))
				IString=Mid(IString,1,inStrRev(IString,myCommand)-1)
				Dim iMode
				iMode=1
				Select Case String(Run)
				Case "normal"
					myMode=2
				Case "page"
					Run=""
					myMode=3
				Case "addnew" '如果条件不成立则新建一条记录
					iMode=3
					myMode=4
				Case "update"
					iMode=3
					myMode=5
				Case Else '自定义的分页查询模式
					myMode=3
				End Select
				Set RS=K.Object("ADODB.RecordSet")
				RS.Open IString,db,1,iMode
			Else
				myMode=1
				Set RS=db.Execute(IString)
			End If
			myCounter=myCounter+1
		Case Else
			myMode=0
			db.Execute IString
		End Select

		'检测错误，并且针对当前状态进行必要的更新。
		If Err.Number=0 Then
			If myMode=0 Then Exit Property
			If RS.Eof Or RS.Bof Then
				If myMode=4 Then RS.addNew() Else RS_Clear
			Else
				If myMode=3 Then Adv Run
				mySafe=True
			End If
		Else
			If myMode>0 Then RS_Clear
			If Global_Test Then Response.Write IString & "<br/>"
			Response.Write "SQL Error:" & Err.Source & "[" & Err.Number & "]；" & Err.Description
			Err.Clear
		End If
	End Property
	Private Sub Adv(IString)
		'初始化分页查询的各项基本参数。
		If IString<>"" Then
			If inStr(IString,",")>1 Then
				myPageTag=Cut(IString,",")
				myPageSize=Cut(IString,",\R")
			Else
				myPageTag=IString
			End If
		End If
		If myPageTag="" Then myPageTag=myTag
		myPageTag=Trim(myPageTag)
		If isNumeric(myPageSize) Then
			myPageSize=Fix(myPageSize)
			If myPageSize<1 Then myPageSize=mySize
		Else
			myPageSize=mySize
		End If

		'重置当前记录数。
		myNowRecord=0
		'取得记录总数。
      		myAllRecord=RS.recordCount
		'计算分页总数。
      		myAllPage=myAllRecord\myPageSize
		If myAllRecord Mod myPageSize<>0 Then myAllPage=myAllPage+1
		'取得当前分页序号。
		myNowPage=K.Value("request." & myPageTag & ",1")
		If myNowPage<1 Then myNowPage=1
		'如果当前分页序号大于分页总数，则重置为最大的分页序号。
		If myNowPage>myAllPage Then myNowPage=myAllPage
		'更新设置。
		RS.pageSize=myPageSize
		RS.absolutePage=myNowPage
	End Sub
	Private Sub RS_Clear()
		If myMode>1 Then RS.Close()
		Set RS=Nothing
		mySafe=False
		myMode=0
	End Sub

	Public Sub Clear()
		Select Case myMode
		Case 0
			Exit Sub
		Case 3
			myAllRecord=0
			myNowRecord=0
			myAllPage=0
			myNowPage=0
			myPageTag=""
			myPageSize=""
		Case 4,5
			RS.Update()
		End Select
		RS_Clear()
	End Sub
	Public Property Get Safe()
		If myMode=0 Then Safe=False Else Safe=mySafe
	End Property
	Public Property Get Last()
		Select Case myMode
		Case 0
			Last=True
		Case 3
			Last=RS.Eof Or myPageSize<myNowRecord+1
		Case Else
			Last=RS.Eof
		End Select
	End Property
	Public Sub Move()
		myNowRecord=myNowRecord+1
		RS.moveNext
	End Sub
	Public Property Let Record(IName,Byval IValue)
		If myMode<4 Then Exit Property

		If isNumeric(IName) And inStr(IName,",")=0 Then
			IName=IName*1
		Else
			If inStr(IName,",")>1 Then
				IValue=K.Format(IValue,Cut(IName,",\R"))
				IName=Cut(IName,",")
			End If
		End If
		RS(IName)=IValue
	End Property
	Public Property Get Record(IName)
		If myMode=0 Then Exit Property

		If isNumeric(IName) And inStr(IName,",")=0 Then
			Record=RS(IName*1)
			If isNull(Record) Then Record=""
		Else
			If inStr(IName,".")=1 Then Record=getValue(Mid(IName,2)):Exit Property
			Dim iFormat
			If inStr(IName,",")>1 Then
				iFormat=Cut(IName,",\R")
				IName=Cut(IName,",")
				If isNumeric(IName) Then IName=1*IName
			End If
			Record=RS(IName)
			If isNull(Record) Then Record=""
			If iFormat<>"" Then Record=K.Format(Record,iFormat)
		End If
	End Property
	Private Function getValue(IName)
		Select Case lCase(IName)
		Case "allrecord","ar"
			getValue=myAllRecord
		Case "nowrecord","nr"
			getValue=myNowRecord
		Case "allpage","ap"
			getValue=myAllPage
		Case "nowpage","np"
			getValue=myNowPage
		Case "pagetag","pt"
			getValue=myPageTag
		Case "pagesize","ps"
			getValue=myPageSize
		Case "tag"
			getValue=myTag
		Case "size"
			getValue=mySize
		End Select
	End Function

	Public Property Get Version(IName)
		Select Case IName
		Case ""
			Version="4.1.5"
		Case "title"
			Version="FastSimple SQL Module"
		Case "full"
			Version=Version("title") & " Version " & Version("")
		Case "author"
			Version="Leo Amos"
		End Select
	End Property
End Class
'增加通用验证函数
Public Function isVoid(IValue)
	Select Case varType(IValue)
	Case 8
		If Trim(IValue)="" Then isVoid=True Else isVoid=False
	Case 2,3,4,5,6,7,11,17
		isVoid=False
	Case Else
		isVoid=True
	End Select
End Function
Public Function isBe(IObject)
	If isObject(IObject) Then
		If IObject Is Nothing Then isBe=False Else isBe=True
	Else
		isBe=False
	End If
End Function
'增加通用工具函数
Public Function Cut(IValue,ISymbol)
	Cut=Trim(IValue)
	If Cut="" Then Exit Function
	ISymbol=Trim(ISymbol) '分隔符对大小写敏感
	If ISymbol="" Then
		ISymbol=":"
	Else
		'获取参数信息
		Dim iLetter,iDirection
		Select Case inStrRev(ISymbol,"\")
		Case 0
		Case 1
			If ISymbol<>"\" Then
				iDirection=uCase(Mid(ISymbol,2))
				ISymbol=":"
			End If
		Case Else
			iDirection=uCase(Mid(ISymbol,inStrRev(ISymbol,"\")+1))
			ISymbol=Mid(ISymbol,1,inStrRev(ISymbol,"\")-1)
		End Select
		Select Case Len(iDirection)
		Case 1
			Select Case iDirection
			Case "U","L"
				iLetter=iDirection
			End Select
		Case 2
			iLetter=Mid(iDirection,2)
			iDirection=Left(iDirection,1)
		End Select
	End If
	'分割字符串
	Select Case iDirection
	Case "R"
		If inStr(Cut,ISymbol)>0 Then Cut=Mid(Cut,inStr(Cut,ISymbol)+Len(ISymbol)) Else Cut=""
	Case Else
		If inStr(Cut,ISymbol)>0 Then Cut=Mid(Cut,1,inStr(Cut,ISymbol)-1)
	End Select
	Cut=Trim(Cut)
	'设置字符串的大小写规则
	Select Case iLetter
	Case "L"
		Cut=lCase(Cut)
	Case "U"
		Cut=uCase(Cut)
	End Select
End Function
Public Function String(IValue)
	If isVoid(IValue) Then Exit Function Else String=Trim(IValue)
	If Right(String,2)="\U" Then String=uCase(Trim(Mid(String,1,Len(String)-2))) Else String=lCase(String)
End Function
%>
<!--#include file="base.asp"-->