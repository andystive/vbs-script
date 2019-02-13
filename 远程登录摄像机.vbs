'本脚本为自动打开摄像机登录界面，并填充默认账号密码
'IP_First 参数为开始ip地址
'IP_Last 参数为终止ip地址
'IP_Range 参数为ip地址段，修改时注意不要落下最后的小数点

Dim user,passwd,ie,IP_First,IP_Last,IP_Range

user = "admin"
passwd = "passwd"
IP_First = 1
IP_Last = 254
IP_Range = "192.168.5."

for i = IP_First To IP_Last
	Set ie = CreateObject("InternetExplorer.Application")
	ie.FullScreen = 0
	ie.Visible = True
	ie.Navigate IP_Range&IP_First
	IP_First = IP_First + 1
	wscript.sleep 2000
	
	if ie.ReadyState <> 4 Then
		MsgBox("摄像机离线，关闭网页，点我继续") ,, True
	else
		ie.document.getelementByid("UserName").value=user
		ie.document.getelementByid("Password").value=passwd
		MsgBox("关闭网页，点我继续") ,, True
	end if
Next

MsgBox("摄像机已巡检完毕，点我退出")