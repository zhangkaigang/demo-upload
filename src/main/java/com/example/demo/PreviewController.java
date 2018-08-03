package com.example.demo;

import net.sf.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * 在C盘新建文件夹tempFiles，并把需要预览的测试文件放入文件夹中
 * Tomcat的配置文件server.xml配置虚拟路径指到tempFiles文件夹
 *	<Context path="/file"  docBase="C:\tempFiles" debug="0" reloadable="true"/>
 * <p>Description:</p>
 */
@Controller
public class PreviewController {

    private Logger logger = LoggerFactory.getLogger(PreviewController.class);
    // 符合预览的格式
    private static final String[] FORMAT = {"txt","jpg","png","jpeg","doc","docx","xls","xlsx","pdf"};
    private static final String RESULT_STAT = "resultStat";
    public static final String REAL_PATH = "realPath";
    private static final String SUCCESS = "SUCCESS";
    private static final String ERROR = "ERROR";
    private static final String TEMP_DIR_PATH = "C:\\tempFiles\\";

    private static Map<String,String> imageContentType = new HashMap<String,String>();
    static {
        imageContentType.put("jpg","image/jpeg");
        imageContentType.put("jpeg","image/jpeg");
        imageContentType.put("png","image/png");
    }

    @RequestMapping("/preview")
    public String preview(){
        return "preview";
    }

    @RequestMapping("/previewAction")
    @ResponseBody
    public void previewAction(HttpServletRequest request, HttpServletResponse response) throws IOException{
        String onlineType = request.getParameter("onlineType");
        if("previw".equals(onlineType)) {
            response.setContentType("text/html;charset=UTF-8");
            response.setCharacterEncoding("utf-8");

            String ftpFileName = request.getParameter("ftpFileName");
            int i = ftpFileName.lastIndexOf('.');
            String prefix = ftpFileName.substring(0,i);
            String suffix = ftpFileName.substring(i+1).toLowerCase();
            // 定义一些变量
            Map<String,Object> map = new HashMap<String,Object>();
            try {
                if(Arrays.asList(FORMAT).contains(suffix)){
                    String scheme = request.getScheme();
                    String ip = request.getServerName();
                    String port = request.getServerPort()+"";
                    String contextPath = request.getContextPath();
                    String servletPath = request.getServletPath();
                    String realPath = "";
                    String result = "";

                    if("pdf".equals(suffix) || "jpg".equals(suffix) || "png".equals(suffix) || "jpeg".equals(suffix)){
                        realPath =scheme+"://"+ip+":"+port+contextPath+servletPath
                                +"?onlineType=onlineView&type="+suffix+"&ftpFileName="+ftpFileName;
                        map.put(RESULT_STAT,SUCCESS);
                        map.put(REAL_PATH,realPath);
                        writeResponse(map, response);
                    }else if("txt".equals(suffix)){
                        response.setContentType("text/plain;charset=utf-8");
                        realPath = scheme+"://"+ip+":"+port+contextPath+servletPath
                                +"?onlineType=txtView&ftpFileName="+ftpFileName;
                        map.put(RESULT_STAT,SUCCESS);
                        map.put(REAL_PATH,realPath);
                        writeResponse(map, response);
                    }else {
                        response.setHeader("Content-type", "text/html;charset=utf-8");
                        response.setContentType("text/html; charset=utf-8");
                        // 属于word、excel则需要转换成html生成32位临时文件夹
                        String tempFilePath = UUID.randomUUID().toString().replace("-", "");
                        String path = TEMP_DIR_PATH + tempFilePath;
                        File dirFile = new File(path);
                        boolean bFile = false;
                        bFile   = dirFile.exists();
                        if(!bFile){
                            bFile = dirFile.mkdir();
                            if(!bFile){// 创建临时文件夹失败
                                map.put(RESULT_STAT,ERROR);
                                map.put("msg","创建临时文件夹失败");
                                writeResponse(map, response);
                                return;
                            }
                        }

                        // 针对excel和word，调用文件转html方法，传递的参数为文件的前缀和后缀，后缀以及路径:
                        result = genHtml(prefix,suffix,path,ftpFileName);

                        map.put(RESULT_STAT,result);
                        if(SUCCESS.equals(result)){
                            realPath = scheme+"://"+ip+":"+port+"/file/"+tempFilePath+"/"+prefix+".html";
                            map.put("tempFilePath",tempFilePath);
                            map.put(REAL_PATH,realPath);
                        }else{
                            map.put("msg","预览失败");
                        }
                        writeResponse(map, response);
                    }
                }else {
                    map.put(RESULT_STAT,ERROR);
                    map.put("msg","不支持该格式文件的预览");
                    writeResponse(map, response);
                    return;
                }
            }catch(Exception e) {
                logger.error(e.getMessage(),e);
            }
        }else if("onlineView".equals(onlineType)) {
            String ftpFileName = request.getParameter("ftpFileName");
            String sourcePath = TEMP_DIR_PATH+ftpFileName;
            int i = ftpFileName.lastIndexOf('.');
            String suffix = ftpFileName.substring(i+1).toLowerCase();
            String type = request.getParameter("type");
            if("pdf".equals(type)){
                response.setContentType("application/pdf");
            }else{
                response.setContentType((String)imageContentType.get(suffix));
            }

            InputStream input = null;
            OutputStream outputStream = response.getOutputStream();
            try {
                input = new FileInputStream(new File(sourcePath));

                int count = 0;
                byte[] buffer = new byte[1024 * 1024];
                while ((count = input.read(buffer)) != -1){
                    outputStream.write(buffer, 0, count);
                }
                outputStream.flush();
            } catch(Exception e) {
                logger.error(e.getMessage(),e);
                response.setContentType("text/plain; charset=utf-8");
                String str = "预览失败" ;        // 准备一个字符串
                byte b[] = str.getBytes() ;            // 只能输出byte数组，所以将字符串变为byte数组
                outputStream.write(b);
            }finally {
                closeStream(outputStream, input);
            }
        }else if("txtView".equals(onlineType)) {
            String ftpFileName = request.getParameter("ftpFileName");
            response.setContentType("text/plain;charset=UTF-8");
            response.setCharacterEncoding("utf-8");
            BufferedReader bis = null;
            InputStream input = null;
            InputStreamReader inputStreamReader = null;
            PrintWriter writer = response.getWriter();
            String sourcePath = TEMP_DIR_PATH+ftpFileName;
            try{
                // 得到服务器配置
                input = new FileInputStream(new File(sourcePath));
                inputStreamReader = new InputStreamReader(input,"GBK");// 文本默认ANSI
                bis = new BufferedReader(inputStreamReader);
                StringBuilder buf=new StringBuilder();
                String temp;
                while ((temp = bis.readLine()) != null) {
                    buf.append(temp);
                    writer.write(temp+"\r\n");
                }
            }catch (Exception e){
                logger.error(e.getMessage(),e);
                writer.write("预览失败");
            }finally {
                if(writer != null) {
                    try{
                        writer.close();
                    }catch(Exception e){
                        logger.error(e.getMessage(),e);
                    }
                }
                if(inputStreamReader!=null){
                    try{
                        inputStreamReader.close();
                    }catch(Exception e){
                        logger.error(e.getMessage(),e);
                    }
                }
                if(input!=null){
                    try{
                        input.close();
                    }catch(Exception e){
                        logger.error(e.getMessage(),e);
                    }
                }
                if(bis!=null){
                    try{
                        bis.close();
                    }catch(Exception e){
                        logger.error(e.getMessage(),e);
                    }
                }

            }
        }
    }

    /**
     * 判断格式文件对应的方法
     * @param prefix
     * @param suffix
     * @param path
     * @return
     * @throws Exception
     */
    private String genHtml(String prefix,String suffix,String path,String ftpFileName) throws Exception{
        if("doc".equals(suffix)){
            return POIUtil.docToHtml(prefix,path,ftpFileName);
        }else if("docx".equals(suffix)){
            return POIUtil.docxToHtml(prefix,path,ftpFileName);
        }else if("xls".equals(suffix)){
            return POIUtil.xlsToHtml(prefix,path,ftpFileName);
        }else if("xlsx".equals(suffix)){
            return POIUtil.xlsxToHtml(prefix,path,ftpFileName);
        }
        return null;
    }

    private void writeResponse(Object obj, HttpServletResponse response) {
        try {
            response.setContentType("text/html;charset=utf-8");//设置编码
            JSONObject jsonObj =  JSONObject.fromObject(obj);
            String mes  = jsonObj.toString();
            PrintWriter writer = response.getWriter();
            writer.write(mes);
            writer.flush();
            writer.close();
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
    }

    /**
     * 关闭流
     * @param outputStream
     * @param inputStream
     */
    private void closeStream(OutputStream outputStream, InputStream inputStream) {
        if(outputStream!=null) {
            try {
                outputStream.close();
            }catch(Exception e) {
                logger.error(e.getMessage(),e);
            }
        }
        if(inputStream!=null) {
            try {
                inputStream.close();
            }catch(Exception e) {
                logger.error(e.getMessage(),e);
            }
        }

    }
}
