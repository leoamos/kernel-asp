项目说明:ASP核心库,包含模板系统
使用方法:把整个Library目录放在项目根目录下,然后在页面中引用library/kernel.asp文件

<!--#include file="library/kernel.asp"-->
<%
K.HTML "Add"
Set K=Nothing
%>

最简单的引用如上,创建res/template目录并新建一个标准XML文件,文件名改为main.master所有的HTML全局页面内容可以保存在该文件中,分页的模板可以另外新建一个XML文件命名和你的页面相同的名称比如index.xml那么该文件将自动指向index.asp文件,所有模板文件的路径结构和程序保持一致

如果需要完整的说明文档请访问我们的开发手册
