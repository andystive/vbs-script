'���ű�Ϊ�Զ����������¼���棬�����Ĭ���˺�����
'IP_First ����Ϊ��ʼip��ַ
'IP_Last ����Ϊ��ֹip��ַ
'IP_Range ����Ϊip��ַ�Σ��޸�ʱע�ⲻҪ��������С����

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
		MsgBox("��������ߣ��ر���ҳ�����Ҽ���") ,, True
	else
		ie.document.getelementByid("UserName").value=user
		ie.document.getelementByid("Password").value=passwd
		MsgBox("�ر���ҳ�����Ҽ���") ,, True
	end if
Next

MsgBox("�������Ѳ����ϣ������˳�")