<%@ page language="java" contentType="text/html; charset=UTF-8"
         pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta charset="UTF-8">
    <title>文件预览测试Demo</title>
    <script type="text/javascript" src="js/jquery-1.10.1.min.js"></script>
    <script type="text/javascript">
        //谷歌浏览器兼容showModalDialog开始
        var has_showModalDialog = !!window.showModalDialog;//定义一个全局变量判定是否有原生showModalDialog方法
        if(!has_showModalDialog &&!!(window.opener)){
            window.onbeforeunload=function(){
                window.opener.hasOpenWindow = false;        //弹窗关闭时告诉opener：它子窗口已经关闭
            }
        }
        //定义window.showModalDialog如果它不存在
        if(window.showModalDialog == undefined){
            window.showModalDialog = function(url,mixedVar,features){
                if(window.hasOpenWindow){
                    window.myNewWindow.focus();
                }
                window.hasOpenWindow = true;
                if(mixedVar) var mixedVar = mixedVar;
                if(features) var features = features.replace(/(dialog)|(px)/ig,"").replace(/;/g,',').replace(/\:/g,"=");
                var left = (window.screen.width - parseInt(features.match(/width[\s]*=[\s]*([\d]+)/i)[1]))/2;
                var top = (window.screen.height - parseInt(features.match(/height[\s]*=[\s]*([\d]+)/i)[1]))/2;
                window.myNewWindow = window.open(url,"_blank",features);
            }
        }
        //谷歌浏览器兼容showModalDialog结束


        function getRootPath(){
            //获取当前网址，如： http://localhost:8083/uimcardprj/share/meun.jsp
            var curWwwPath=window.document.location.href;
            //获取主机地址之后的目录，如： uimcardprj/share/meun.jsp
            var pathName=window.document.location.pathname;
            var pos=curWwwPath.indexOf(pathName);
            //获取主机地址，如： http://localhost:8083
            var localhostPaht=curWwwPath.substring(0,pos);
            //获取带"/"的项目名，如：/uimcardprj
            var projectName=pathName.substring(0,pathName.substr(1).indexOf('/')+1);
            return(localhostPaht+projectName);
        }

        // 预览doc
        function docPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test2003.doc"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");


                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览docx
        function docxPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test2007.docx"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览xls
        function xlsPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test2003.xls"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览xlsx
        function xlsxPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test2007.xlsx"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览PDF
        function pdfPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test.pdf"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览png
        function pngPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test.png"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览jpg
        function jpgPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test.jpg"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览jpeg
        function jpegPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test.jpeg"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

        // 预览txt
        function txtPreview(){
            $.ajax({
                async: false,
                type: "post",
                dataType : "json",
                url: "previewAction?onlineType=previw",
                data:{
                    ftpFileName:"test.txt"
                },
                success:function(data){
                    if(data.resultStat == "SUCCESS"){
                        var url = data.realPath;
                        window.showModalDialog(url, null,
                                "dialogWidth:950px;dialogHeight:550px;help:no;resizable:yes");
                    }else if(data.resultStat=="ERROR"){
                        alert(data.msg);
                    }else{
                        alert("预览失败");
                    }

                }
            });
        }

    </script>
</head>
<body>
<input type="button" value="预览doc" onclick="docPreview()"/>
<input type="button" value="预览docx" onclick="docxPreview()"/>
<input type="button" value="预览xls" onclick="xlsPreview()"/>
<input type="button" value="预览xlsx" onclick="xlsxPreview()"/>
<br><br>
<input type="button" value="预览pdf" onclick="pdfPreview()">
<input type="button" value="预览png" onclick="pngPreview()">
<input type="button" value="预览jpg" onclick="jpgPreview()">
<input type="button" value="预览jpeg" onclick="jpegPreview()">
<input type="button" value="预览txt" onclick="txtPreview()">

</body>
</html>