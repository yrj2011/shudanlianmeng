<%
dim sessionvar,WebName,weburl,webmail,webceo,noceo,webinfo,badReg,offweb,webopen,webqq,adflperpage
Dim mailzj,mailkg,mailaddress,mailusername,mailuserpass,mailsend,mailname
Dim usery,upass,userfb,userdj,fbpass,vippass,djdown,fsMessage,fanuser,bodyMessage,Messagesum,friendsum,collectsum,Specialsum,fbpoints,adpoints,Attmesum
Dim canupload,uploadsize,uploadnum,uploadtype,webdh
Dim zhvip,aduser,viptime,zvippoints
Dim dj586_In,bad_ip,bad_ID
Dim tj,Down_Mp3,asphtml,musicindex,listhtml,uploadtypeb
Dim Ylmv(2),logoyes,logoImg,logoImgW,logoFnot,logoFnotsize,logoSpan,logoFnotname,logoFnotcolor
'=====网站基本信息=====
sessionvar="jieducm"          '设置系统变量
webname="系统官方演示"                '设置站点名称
weburl="www.jieducm.cn"                  '设置网站地址
weburl2="www.258shua.com"                  '会员网站地址
webmail="616591415@qq.com"                '设置站长EMAIL
webceo="小笨"                  '设置站长名字
noceo="小笨"                    '替换站长名称
webqq="616591415"                    '设置站长QQ
webdh="11111111"                    '设置站长电话
ceopass="123456"                '设置站长代号
webinfo="备案中..."                '设置备案信息
badReg="DJ586|admin|管理员|传世私服"                  '禁止注册的ID
offweb="网站正在维护中......感谢您使用捷度传媒独立开发的互刷平台www.jieducm.cn"                  '网站维护
webopen="NO"                '网站维护开关
dj586_In="FUCK|传世私服|传世|你妈的|fuck|google456|操你妈|爸爸|爷爷|奶奶|姥姥|孙子|儿子|江泽|胡锦|Fuck|无耻|你妈"                '设置屏蔽内容
bad_ip=""                  '设置屏蔽IP
bad_ID=""                  '设置屏蔽ID
tj="<script language=JavaScript src=/Js/Tj.js></script>"                          '统计代码
adflperpage="50"        '后台连接每页显示数量
Ylmv(0)="50"                '分类栏目每页显示数量
'=====网站生成模式=====
asphtml="asp"                '页面格式
listhtml="1"              '栏目列表文件的存放位置
'=====首页播放器播放模式=====
musicindex="0"          '播放模式
'=====下载开关设置=====
Down_Mp3="1"              '下载开关
'=====邮件相关设置=====
mailzj=0                                      '邮件组件选择
mailkg=0                                      'email开关
mailaddress="smtp.163.com"        '邮件服务器地址
mailusername=""      '登录名
mailuserpass=""      '登录密码
mailsend="9984net@163.com"              '发送邮箱
mailname="管理系统"              '发送时显示的姓名
'=====会员相关设置=====
usery="yes"                    '设置注册
upass=0                                        '注册模式
userfb="yes"                  '设置发表
userdj="yes"                  '设置音乐发表
fbpass="yes"                  '普通会员发表审核
vippass="yes"                'VIP会员发表审核
djdown="yes"                  '音乐下载
fsMessage="yes"            '是不是发送短消息
fanuser="no"      '是否开启泛域名
bodyMessage="本站欢迎您的到来"        '发送短消息内容
Messagesum=50                             '收信箱储存短信条数
friendsum=1000                               '好友个数
collectsum=1000                             '音乐收藏数
Specialsum=200                             '专集收藏数
Attmesum=300                                 '嗨友团限制数
fbpoints=10                                 '发表积分
adpoints=2                                 '宣传积分
'=====VIP相关设置=====
zhvip="yes"                    '是否开放会员转VIP功能
aduser="yes"                    '是否开放会员转VIP功能
viptime=30                                     'vip天数
zvippoints=300                               '转VIP会员的积分数
'=====上传信息设置=====
canupload=1                                 '设置上传等级
uploadsize=200000                               '设置上传最大数
uploadnum=50                                 '设置上传数量
uploadtype="gif|jpg|bmp"          '设置上传格式
uploadtypeb="mp3"        '设置上传格式
'=====图片水印设置=====
logoyes=0                                     '设置水印开关
logoImg="/Images/Up_Logo.gif"                '设置水印图片的路径
logoImgW=125                                   '设置水印图片的宽度
logoFnot="图片上传自 www.jieducm.cn"              '设置文字水印名称
logoFnotsize=12                           '设置水印字体大小
logoSpan=4                                   '设置写入位置
logoFnotname="新宋体"      '设置水印字体名称
logoFnotcolor="&HFF0000"    '设置文字水印颜色
%>
