# demo-upload
预览功能，预览word以及excel，原理是通过POI将word转为html进行展示
预览：
http://localhost:8080/preview

此项工作主要是针对word和excel的前期工作：
1.  在C盘新建tempFiles文件夹，里面放入相应文件，文件名称参照preview.jsp中的文件名
2.	本地C盘需要提前建立好文件夹tempFiles，如果是在linux则需要在home下建立tempFiles(路径不能错，因为代码暂时固化了)

3.	tomcat的server.xml需要配置虚拟路径(如果是用springboot项目则可以在代码里对内置的tomcat进行虚拟路径配置)
<Context path="/file"  docBase="C:\tempFiles" debug="0" reloadable="true"/>或者
<Context path="/file"  docBase="/home/tempFiles" debug="0" reloadable="true"/>
