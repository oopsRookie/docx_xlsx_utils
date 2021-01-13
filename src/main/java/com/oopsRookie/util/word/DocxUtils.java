package com.oopsRookie.util.word;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.core.io.ClassPathResource;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import javax.xml.namespace.QName;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.regex.Pattern;

public class DocxUtils {

    /**
     * 根据模板文件uri生成docx文件
     * @params： data : 要渲染的数据对象
     * @params： templatePath : 模板的uri，模板必须为docx文件。静态资源static下的路径
     * @params： outputFileName: 返回给客户端的文件名，后缀扩展名无需设置
     * @params： response : HttpServletResponse
     */
    public static <T> void genWord(T data, String templatePath, String outputFileName, HttpServletResponse response) throws IOException {

        //通过流来操作，防止项目打包成jar时获取不到内部文件
        ClassPathResource classPathResource = new ClassPathResource(templatePath);
        InputStream is = classPathResource.getInputStream();

        genWord(data, is, outputFileName, response);
    }

    /**
     * 根据模板文件对象File生成docx文件
     * @params： data : 要渲染的数据对象
     * @params： templateFile : 模板文件，必须为docx文件
     * @params： outputFileName: 返回给客户端的文件名，后缀扩展名无需设置
     * @params： response : HttpServletResponse
     */
    public static <T> void genWord(T data, File templateFile, String outputFileName, HttpServletResponse response) throws IOException {
        if (!templateFile.getName().endsWith(".docx")) {
            throw new RuntimeException(templateFile.getName() + "不是.docx文件，请选择docx文件作为模板");
        }
        genWord(data, new FileInputStream(templateFile), outputFileName, response);

    }


    /**
     *
     * @param data 要渲染的数据对象
     * @param is InputStream模板文件输入流
     * @param outputFileName 返回给客户端的文件名，后缀扩展名无需设置
     * @param response HttpServletResponse
     * @param <T>
     * @throws IOException
     */
    public static <T> void genWord(T data, InputStream is, String outputFileName, HttpServletResponse response) throws IOException {

        genWord(data, is, outputFileName, null, response);
    }

    /**
     * 根据模板生成word，并添加水印
     * @param data 要渲染的数据对象
     * @param templatePath 模板路径
     * @param outputFileName 返回给客户端的文件名，后缀扩展名无需设置
     * @param watermark 水印名
     * @param response HttpServletResponse
     * @param <T>
     * @throws IOException
     */
    public static <T> void genWord(T data, String templatePath, String outputFileName, String watermark, HttpServletResponse response) throws IOException {
        ClassPathResource classPathResource = new ClassPathResource(templatePath);
        InputStream is = classPathResource.getInputStream();
        genWord(data, is, outputFileName, watermark, response);
    }


    private static <T> void genWord(T data, InputStream is, String outputFileName, String watermark, HttpServletResponse response) throws IOException {
        boolean openWatermark = false;
        if (watermark != null && !"".equals(watermark)) {
            openWatermark = true;
        }
        if (outputFileName == null ||
                outputFileName.equals("")
                || outputFileName.trim().length() == 0) {
            outputFileName = "default_filename";
        }

        String fileName = URLEncoder.encode(outputFileName.trim() + ".docx",
                "UTF-8");     //alternative: StandardCharsets.UTF_8 (since java 10)
        ConfigureBuilder builder = new ConfigureBuilder();

        //compile model file to InputStream
        XWPFTemplate template = XWPFTemplate.compile(is, builder.build())
                .render(data);

        //set watermark if need
        if (openWatermark) {
            setWaterMark(template, watermark);
        }

        /* 输出方式一：先生成本地文件，再返回文件流 */
        //FileOutputStream fos = new FileOutputStream("src/main/resources/static/gen/output.docx");     //输出至文件

        //输出方式二：直接通过response获取servletOutputStream流返回
        response.setCharacterEncoding("UTF-8");
        /* 设置docx格式响应头及文件名 */
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.addHeader("Content-Disposition", "attachment; filename=" + fileName);

        ServletOutputStream fos = response.getOutputStream();        //获取response的输出流

        /* write input stream to outputStream */
        template.write(fos);

        fos.flush();
        fos.close();

        template.close();
    }

    /**
     * 设置word水印格式
     * 模板格式要求：word模板一定不能有眉页和眉尾（包括换行符，若页面没有开启回车符，
     * 请点击：开始 -选项 -显示 - 段落标记 中开启显示回车符）
     * 文档内容最好不要有多余换行符。
     *
     * 注：poi设置水印的本质就是操作眉页
     * @param styleStr
     * @param height
     * @return
     */
    private static String getWaterMarkStyle(String styleStr, double height, double width) {
        Pattern p = Pattern.compile(";");
        String[] strs = p.split(styleStr);
        for (String str : strs) {
            if (str.startsWith("height:")) {
                String heightStr = "height:" + height + "pt";
                styleStr = styleStr.replace(str, heightStr);
                break;
            }
//            else if (str.startsWith("width:")) {
//                String widthStr = "width:" + width + "pt";
//                styleStr = styleStr.replace(str, widthStr);
//                break;
//            }
        }
        return styleStr;
    }

    //水印的模板，不要设置空的眉页和眉脚。
    private static void setWaterMark(XWPFTemplate template, String watermark) {
        //add on 2021年1月10日
        NiceXWPFDocument xwpfDocument = template.getXWPFDocument();

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(xwpfDocument);

        policy.createWatermark(watermark);
        XWPFHeader header = policy.getHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph paragraph = header.getParagraphArray(0);
        CTP ctp = paragraph.getCTP();
        if (ctp.getRList().size() == 0) {
            ctp.addNewR();
            CTR ctr = ctp.getRArray(0);
            ctr.addNewPict();
        }

        XmlObject[] xmlobjects = paragraph.getCTP().getRArray(0).getPictArray(0).selectChildren(new QName("urn:schemas-microsoft-com:vml", "shape"));
        if (xmlobjects.length > 0) {
            com.microsoft.schemas.vml.CTShape ctshape = (com.microsoft.schemas.vml.CTShape) xmlobjects[0];
            //水印颜色
            ctshape.setFillcolor("#7E8B92");

            //水印样式
            ctshape.setStyle(getWaterMarkStyle(ctshape.getStyle(), 35, 400) + ";rotation:315");
        }
    }
}