<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td><strong>ubb的使用</strong>:在添加新闻时会有一些像<img src="pic/move.gif" width="23" height="22">这样的按扭,你点击一下他,会弹出一个JAVA输入框,然后你在框内要输入你要输入的字符,这样在外显示你输入的那几个字就会移动<br>
      <img src="pic/glow.gif" width="23" height="22"> 点击这个按扭输入进去的字显示出来时会有发光的效果<br>
      <img src="pic/icon_editor_email.gif" width="23" height="22">这个是EMAIL的超级链接,在输入框中输入的e-mail显示出来时会有超级链接的效果<br>
      <img src="pic/icon_editor_url.gif" width="23" height="22"> 这个是网址的超级链接,在输入框中输入的网址显示出来时会有超级链接<br>
      <img src="pic/icon_editor_list.gif" width="23" height="22">点击这个图标后会要你选择a或是数字,选数字后再输入的字符会以选择题一样在你输入的每一个行字符前添加相应的数字,选择字母也是一样,会有如下效果<br>
      a.我输入的第一行字<br>
      b,我输入的第二行<br>
      ......<br>
      一直到我不输入数字他才结束.大致的就介绍这么多,你在使用的过程序会慢慢发现其用法的</td>
  </tr>
  <tr>
    <td>　　<strong>视频点播文件</strong>要详细的说一下,当文件大于3M以后建议你不要用后台程序上传.因为HTTP的传输速度慢得跟蜗牛一样,请你把你要上传的视频点播文件用FTP上传到网站服务器的admingly/movie目录下.然后只要记住视频文件名.再到后台的视频文件管理中,选择添加已知的视频文件,再把视频文件的名字输入进去.大于3M以上的文件千万不要用后台上传.否则浪费你大量的时间不要怪我.如果HTTP的上传速度够快的话也不会有FTP的出现.</td>
  </tr>
  <tr>
    <td>　　<strong>添加管理员</strong>时请注意: 添加的管理员具有与你一样的权限,他也可以删除你的任何资料,包括删除你的管理员账号.</td>
  </tr>
  <tr>
    <td> 　　<strong>cookies</strong>可以在后台任意的开关.当打开cookies时,在你电脑上会保存一份你的管理员账号资料一份,当你登入后台管理时可以不用输入密码而直接登陆.关闭时就是每次进后台都要输入账号和密码.不过为了安全起见,建议您关闭cookies,除非你能保证你的电脑安全系数相当高.</td>
  </tr>
</table>

<%

%>
</body>
</html>
