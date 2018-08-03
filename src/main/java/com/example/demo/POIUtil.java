package com.example.demo;

import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

/**
 *
 */
public class POIUtil {
    private static Logger logger = LoggerFactory.getLogger(POIUtil.class);
    private static final String SUCCESS = "SUCCESS";
    private static final String ERROR ="ERROR";
    private static final String TEMP_DIR_PATH = "C:\\tempFiles\\";
    /**
     * word 2003 doc文档转成html方法
     * @param prefix
     * @param path
     * @return
     */
    public static String docToHtml(String prefix,String path,String ftpFileName){
        path = path+"\\";
        String sourcePath = TEMP_DIR_PATH+ftpFileName;
        // 指定生成的html全路径
        String fileHtml = path + prefix + ".html";
        InputStream input = null;
        ByteArrayOutputStream outStream = null;
        try{
            //根据输入文件路径与名称读取文件流
            input = new FileInputStream(new File(sourcePath));
            //把文件流转化为输入wordDom对象
            HWPFDocument wordDocument = new HWPFDocument(input);
            //生成针对Dom对象的转化器(通过反射构建dom创建者工厂,生成dom创建者,生成dom对象)
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                    DocumentBuilderFactory.newInstance().newDocumentBuilder()
                            .newDocument());
            //转化器重写内部方法
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                public String savePicture(byte[] content, PictureType pictureType,
                                          String suggestedName, float widthInches, float heightInches) {
                    return suggestedName;
                }

            });
            //转化器开始转化接收到的dom对象
            wordToHtmlConverter.processDocument(wordDocument);
            //保存文档中的图片
            List pics = wordDocument.getPicturesTable().getAllPictures();
            FileOutputStream fos = null;
            if (pics != null) {
                for (int i = 0; i < pics.size(); i++) {
                    Picture pic = (Picture)pics.get(i);
                    try {
                        fos = new FileOutputStream(path + pic.suggestFullFileName());
                        pic.writeImageContent(fos);
                    } catch (FileNotFoundException e) {
                        logger.error(e.getMessage(),e);
                    }finally {
                        if(fos != null){
                            try {
                                fos.close();
                            } catch (IOException e) {
                                logger.error(e.getMessage(),e);
                            }
                        }
                    }
                }
            }
            //从加载了输入文件中的转换器中提取DOM节点
            Document htmlDocument = wordToHtmlConverter.getDocument();
            //从提取的DOM节点中获得内容
            DOMSource domSource = new DOMSource(htmlDocument);
            //字节码输出流
            outStream = new ByteArrayOutputStream();
            //输出流的源头
            StreamResult streamResult = new StreamResult(outStream);
            //转化工厂生成序列转化器
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            //设置序列化内容格式
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            String content = new String(outStream.toByteArray(),"utf-8");
            //生成文件方法
            FileUtils.writeStringToFile(new File(fileHtml), content, "utf-8");
            return SUCCESS;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            return ERROR;
        }finally {
            if(outStream!=null){
                try {
                    outStream.close();
                }catch (IOException e){
                    logger.error(e.getMessage(),e);
                }
            }
            if(input!=null){
                try {
                    input.close();
                }catch (IOException e){
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }


    /**
     * word 2007 docx文档转成html方法
     * @param prefix
     * @param path
     * @return
     */
    public static String docxToHtml(String prefix,String path,String ftpFileName){
        path = path+"\\";
        // 指定生成的html全路径
        String fileHtml = path + prefix + ".html";
        InputStream input = null;
        OutputStream out = null;
        String sourcePath = TEMP_DIR_PATH+ftpFileName;
        try{
            //根据输入文件路径与名称读取文件流
            input = new FileInputStream(new File(sourcePath));
            XWPFDocument document = new XWPFDocument(input);
            File imageFolderFile = new File(path);
            //加载html页面时图片路径
            XHTMLOptions options = XHTMLOptions.create().URIResolver( new BasicURIResolver("./"));
            //图片保存文件夹路径
            options.setExtractor(new FileImageExtractor(imageFolderFile));
            out = new FileOutputStream(new File(fileHtml));
            XHTMLConverter.getInstance().convert(document, out, options);
            return SUCCESS;
        }catch (Exception e){
            logger.error(e.getMessage(),e);
            return ERROR;
        }finally{
            if(out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
            if(input!=null){
                try {
                    input.close();
                }catch (IOException e){
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }

    /**
     * excel2003 xls转html
     * @param prefix
     * @param path
     * @return
     */
    public static String xlsToHtml(String prefix,String path,String ftpFileName){
        path = path+"\\";
        String fileHtml = path + prefix + ".html";// 指定生成的html全路径
        InputStream input = null;
        ByteArrayOutputStream outStream = null;
        String sourcePath = TEMP_DIR_PATH+ftpFileName;
        try{
            //读取文档内容
            input = new FileInputStream(new File(sourcePath));
            HSSFWorkbook excelBook=new HSSFWorkbook(input);
            ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter (DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument() );
            excelToHtmlConverter.setOutputColumnHeaders(false);
            excelToHtmlConverter.setOutputRowNumbers(false);
            excelToHtmlConverter.processWorkbook(excelBook);

            Document htmlDocument =excelToHtmlConverter.getDocument();
            outStream = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource (htmlDocument);
            StreamResult streamResult = new StreamResult (outStream);
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty (OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty (OutputKeys.INDENT, "yes");
            serializer.setOutputProperty (OutputKeys.METHOD, "html");
            serializer.transform (domSource, streamResult);

            String content = new String (outStream.toByteArray(),"utf-8");

            FileUtils.writeStringToFile(new File(fileHtml), content, "utf-8");
            return SUCCESS;

        }catch (Exception e){
            logger.error(e.getMessage(),e);
            return ERROR;
        }finally {
            if(outStream!=null){
                try {
                    outStream.close();
                }catch (IOException e){
                    logger.error(e.getMessage(),e);
                }
            }
            if(input!=null){
                try {
                    input.close();
                }catch (IOException e){
                    logger.error(e.getMessage(),e);
                }
            }
        }

    }

    /**
     * excel03是读取文件整个内容转为字符串存进html，excel07是读取文件内容拼成字符串存进html
     * 此方法可以改造成03版本也可以用的，但是样式有点问题所以不用
     * @param prefix
     * @param path
     * @return
     */
    public static String xlsxToHtml(String prefix,String path,String ftpFileName) {
        boolean isWithStyle = true;// 样式
        path = path+"\\";
        String fileHtml = path + prefix + ".html";// 指定生成的html全路径
        InputStream input = null;
        String htmlExcel = null;
        Workbook wb = null;
        String sourcePath = TEMP_DIR_PATH+ftpFileName;
        try {
            //读取文档内容
            input = new FileInputStream(new File(sourcePath));
            wb = WorkbookFactory.create(input);
            XSSFWorkbook xWb = (XSSFWorkbook) wb;
            htmlExcel = getExcelInfo(xWb, isWithStyle);
            FileUtils.writeStringToFile(new File(fileHtml), htmlExcel, "GBK");
            return SUCCESS;
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            return ERROR;
        } finally {
            if(input != null){
                try {
                    input.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }

    private static String getExcelInfo(Workbook wb, boolean isWithStyle) {
        StringBuilder sb = new StringBuilder();
        int sheetCounts = wb.getNumberOfSheets();

        for (int i = 0; i < sheetCounts; i++) {
            Sheet sheet = wb.getSheetAt(i);// 获取第一个Sheet的内容
            int lastRowNum = sheet.getLastRowNum();
            Map<String, String> map[] = getRowSpanColSpanMap(sheet);
            sb.append("<br><br>");
            sb.append(sheet.getSheetName());
            sb.append("<table style='border-collapse:collapse;' width='100%'>");
            Row row = null; // 兼容
            Cell cell = null; // 兼容
            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
                row = sheet.getRow(rowNum);
                if (row == null) {
                    sb.append("<tr><td > &nbsp;</td></tr>");
                    continue;
                }
                sb.append("<tr>");
                int lastColNum = row.getLastCellNum();
                for (int colNum = 0; colNum < lastColNum; colNum++) {
                    cell = row.getCell(colNum);
                    if (cell == null) { // 特殊情况 空白的单元格会返回null
                        sb.append("<td>&nbsp;</td>");
                        continue;
                    }

                    String stringValue = getCellValue(cell);
                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String pointString = map[0].get(rowNum + "," + colNum);
                        map[0].remove(rowNum + "," + colNum);
                        int bottomeRow = Integer.parseInt(pointString.split(",")[0]);
                        int bottomeCol = Integer.parseInt(pointString.split(",")[1]);
                        int rowSpan = bottomeRow - rowNum + 1;
                        int colSpan = bottomeCol - colNum + 1;
                        sb.append("<td rowspan= '" + rowSpan + "' colspan= '" + colSpan + "' ");
                    } else if (map[1].containsKey(rowNum + "," + colNum)) {
                        map[1].remove(rowNum + "," + colNum);
                        continue;
                    } else {
                        sb.append("<td ");
                    }

                    // 判断是否需要样式
                    if (isWithStyle) {
                        dealExcelStyle(wb, sheet, cell, sb);// 处理单元格样式
                    }

                    sb.append(">");
                    if (stringValue == null || "".equals(stringValue.trim())) {
                        sb.append(" &nbsp; ");
                    } else {
                        // 将ascii码为160的空格转换为html下的空格（&nbsp;）
                        sb.append(stringValue.replace(String.valueOf((char) 160), "&nbsp;"));
                    }
                    sb.append("</td>");
                }
                sb.append("</tr>");
            }
            sb.append("</table>");
        }

        return sb.toString();
    }

    private static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {

        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 };
        return map;
    }

    /**
     * 200 * 获取表格单元格Cell内容 201 * @param cell 202 * @return 203
     */
    private static String getCellValue(Cell cell) {

        String result = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case Cell.CELL_TYPE_STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case Cell.CELL_TYPE_BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 251 * 处理表格样式 252 * @param wb 253 * @param sheet 254 * @param cell 255
     * * @param sb 256
     */
    private static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuilder sb) {

        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");// 单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignment();
            sb.append("valign='" + convertVerticalAlignToHtml(verticalAlignment) + "' ");// 单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                short boldWeight = xf.getBoldweight();
                sb.append("style='");
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                sb.append("width:" + columnWidth + "px;");

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }

                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }
                sb.append(getBorderStyle(0, cellStyle.getBorderTop(),
                        ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                sb.append(getBorderStyle(1, cellStyle.getBorderRight(),
                        ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                sb.append(getBorderStyle(2, cellStyle.getBorderBottom(),
                        ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                sb.append(getBorderStyle(3, cellStyle.getBorderLeft(),
                        ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));

            } else if (wb instanceof HSSFWorkbook) {

                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                short boldWeight = hf.getBoldweight();
                short fontColor = hf.getColor();
                sb.append("style='");
                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                sb.append("font-weight:" + boldWeight + ";"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex());
                sb.append("width:" + columnWidth + "px;");
                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                }
                sb.append(getBorderStyle(palette, 0, cellStyle.getBorderTop(), cellStyle.getTopBorderColor()));
                sb.append(getBorderStyle(palette, 1, cellStyle.getBorderRight(), cellStyle.getRightBorderColor()));
                sb.append(getBorderStyle(palette, 3, cellStyle.getBorderLeft(), cellStyle.getLeftBorderColor()));
                sb.append(getBorderStyle(palette, 2, cellStyle.getBorderBottom(), cellStyle.getBottomBorderColor()));
            }

            sb.append("' ");
        }
    }

    /**
     * 330 * 单元格内容的水平对齐方式 331 * @param alignment 332 * @return 333
     */
    private static String convertAlignToHtml(short alignment) {

        String align = "left";
        switch (alignment) {
            case CellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case CellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case CellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * 354 * 单元格中内容的垂直排列方式 355 * @param verticalAlignment 356 * @return 357
     */
    private static String convertVerticalAlignToHtml(short verticalAlignment) {

        String valign = "middle";
        switch (verticalAlignment) {
            case CellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case CellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case CellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

    private static String convertToStardColor(HSSFColor hc) {

        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return sb.toString();
    }

    private static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    static String[] bordesr = { "border-top:", "border-right:", "border-bottom:", "border-left:" };
    static String[] borderStyles = { "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ", "solid ",
            "solid ", "solid", "solid", "solid", "solid", "solid" };

    private static String getBorderStyle(HSSFPalette palette, int b, short s, short t) {

        if (s == 0)
            return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        ;
        String borderColorStr = convertToStardColor(palette.getColor(t));
        borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000" : borderColorStr;
        return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";

    }

    private static String getBorderStyle(int b, short s, XSSFColor xc) {

        if (s == 0)
            return bordesr[b] + borderStyles[s] + "#d0d7e5 1px;";
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();
            borderColorStr = borderColorStr == null || borderColorStr.length() < 1 ? "#000000"
                    : borderColorStr.substring(2);
            return bordesr[b] + borderStyles[s] + borderColorStr + " 1px;";
        }

        return "";
    }

}

