<%
dim sessionvar,WebName,weburl,webmail,webceo,noceo,webinfo,badReg,offweb,webopen,webqq,adflperpage
Dim mailzj,mailkg,mailaddress,mailusername,mailuserpass,mailsend,mailname
Dim usery,upass,userfb,userdj,fbpass,vippass,djdown,fsMessage,fanuser,bodyMessage,Messagesum,friendsum,collectsum,Specialsum,fbpoints,adpoints,Attmesum
Dim canupload,uploadsize,uploadnum,uploadtype,webdh
Dim zhvip,aduser,viptime,zvippoints
Dim dj586_In,bad_ip,bad_ID
Dim tj,Down_Mp3,asphtml,musicindex,listhtml,uploadtypeb
Dim Ylmv(2),logoyes,logoImg,logoImgW,logoFnot,logoFnotsize,logoSpan,logoFnotname,logoFnotcolor
'=====��վ������Ϣ=====
sessionvar="jieducm"          '����ϵͳ����
webname="ϵͳ�ٷ���ʾ"                '����վ������
weburl="www.jieducm.cn"                  '������վ��ַ
weburl2="www.258shua.com"                  '��Ա��վ��ַ
webmail="616591415@qq.com"                '����վ��EMAIL
webceo="С��"                  '����վ������
noceo="С��"                    '�滻վ������
webqq="616591415"                    '����վ��QQ
webdh="11111111"                    '����վ���绰
ceopass="123456"                '����վ������
webinfo="������..."                '���ñ�����Ϣ
badReg="DJ586|admin|����Ա|����˽��"                  '��ֹע���ID
offweb="��վ����ά����......��л��ʹ�ýݶȴ�ý���������Ļ�ˢƽ̨www.jieducm.cn"                  '��վά��
webopen="NO"                '��վά������
dj586_In="FUCK|����˽��|����|�����|fuck|google456|������|�ְ�|үү|����|����|����|����|����|����|Fuck|�޳�|����"                '������������
bad_ip=""                  '��������IP
bad_ID=""                  '��������ID
tj="<script language=JavaScript src=/Js/Tj.js></script>"                          'ͳ�ƴ���
adflperpage="50"        '��̨����ÿҳ��ʾ����
Ylmv(0)="50"                '������Ŀÿҳ��ʾ����
'=====��վ����ģʽ=====
asphtml="asp"                'ҳ���ʽ
listhtml="1"              '��Ŀ�б��ļ��Ĵ��λ��
'=====��ҳ����������ģʽ=====
musicindex="0"          '����ģʽ
'=====���ؿ�������=====
Down_Mp3="1"              '���ؿ���
'=====�ʼ��������=====
mailzj=0                                      '�ʼ����ѡ��
mailkg=0                                      'email����
mailaddress="smtp.163.com"        '�ʼ���������ַ
mailusername=""      '��¼��
mailuserpass=""      '��¼����
mailsend="9984net@163.com"              '��������
mailname="����ϵͳ"              '����ʱ��ʾ������
'=====��Ա�������=====
usery="yes"                    '����ע��
upass=0                                        'ע��ģʽ
userfb="yes"                  '���÷���
userdj="yes"                  '�������ַ���
fbpass="yes"                  '��ͨ��Ա�������
vippass="yes"                'VIP��Ա�������
djdown="yes"                  '��������
fsMessage="yes"            '�ǲ��Ƿ��Ͷ���Ϣ
fanuser="no"      '�Ƿ���������
bodyMessage="��վ��ӭ���ĵ���"        '���Ͷ���Ϣ����
Messagesum=50                             '�����䴢���������
friendsum=1000                               '���Ѹ���
collectsum=1000                             '�����ղ���
Specialsum=200                             'ר���ղ���
Attmesum=300                                 '������������
fbpoints=10                                 '�������
adpoints=2                                 '��������
'=====VIP�������=====
zhvip="yes"                    '�Ƿ񿪷Ż�ԱתVIP����
aduser="yes"                    '�Ƿ񿪷Ż�ԱתVIP����
viptime=30                                     'vip����
zvippoints=300                               'תVIP��Ա�Ļ�����
'=====�ϴ���Ϣ����=====
canupload=1                                 '�����ϴ��ȼ�
uploadsize=200000                               '�����ϴ������
uploadnum=50                                 '�����ϴ�����
uploadtype="gif|jpg|bmp"          '�����ϴ���ʽ
uploadtypeb="mp3"        '�����ϴ���ʽ
'=====ͼƬˮӡ����=====
logoyes=0                                     '����ˮӡ����
logoImg="/Images/Up_Logo.gif"                '����ˮӡͼƬ��·��
logoImgW=125                                   '����ˮӡͼƬ�Ŀ��
logoFnot="ͼƬ�ϴ��� www.jieducm.cn"              '��������ˮӡ����
logoFnotsize=12                           '����ˮӡ�����С
logoSpan=4                                   '����д��λ��
logoFnotname="������"      '����ˮӡ��������
logoFnotcolor="&HFF0000"    '��������ˮӡ��ɫ
%>
