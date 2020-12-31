package com.oopsRookie.util.word;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.ConfigureBuilder;
import org.springframework.core.io.ClassPathResource;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

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
     * 根据模板文件生成docx文件
     * @params： data : 要渲染的数据对象
     * @params： is : InputStream模板文件输入流
     * @params： outputFileName: 返回给客户端的文件名，后缀扩展名无需设置
     * @params： response : HttpServletResponse
     */
    public static <T> void genWord(T data, InputStream is, String outputFileName, HttpServletResponse response) throws IOException {

        if (outputFileName == null ||
                outputFileName.equals("")
                || outputFileName.trim().length() == 0) {
            outputFileName = "default_filename";
        }

        //UFT-8编码防止中文不被解析
        String fileName = URLEncoder.encode(outputFileName.trim() + ".docx", "UTF8");
        //java 10+ 可以采用 StandardCharsets.UTF_8替代UTF8

        ConfigureBuilder builder = new ConfigureBuilder();
        //builder.setElMode(Configure.ELMode.SPEL_MODE);          //开启spring el
        //spring el无法处理中文引号，而word里输英文引号还是中文的


        //compile model file to InputStream
        XWPFTemplate template = XWPFTemplate.compile(is, builder.build())
                .render(data);


        /* 输出方式一：先生成本地文件，再返回文件流 */
        //FileOutputStream fos = new FileOutputStream("src/main/resources/static/gen/output.docx");     //输出至文件

        //输出方式二：直接通过response获取servletOutputStream流返回
        response.setCharacterEncoding("UTF-8");
        /* 设置docx格式响应头及文件名 */
        // MIME types(IANA media types) 链接如下：
        // https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.addHeader("Content-Disposition", "attachment; filename=" + fileName);

        ServletOutputStream fos = response.getOutputStream();        //获取response的输出流

        /* write input stream to outputStream */
        template.write(fos);

        fos.flush();
        fos.close();

        template.close();
    }
}