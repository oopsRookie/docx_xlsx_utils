package com.oopsRookie.util.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.springframework.core.io.ClassPathResource;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class XlsxUtils<E1> {
    private final List<E1> data;
    private final InputStream is;
    private final String outputFileName;
    private final List<Integer> sheetNos;
    private final List<String> sheetNames;
    private final HttpServletResponse response;
    private Map<String,Object> map;

    private XlsxUtils(ExcelBuilder<E1> builder) {
        this.data = builder.data;
        this.is = builder.is;
        this.outputFileName = builder.outputFileName;
        this.sheetNos = builder.sheetNos;
        this.sheetNames = builder.sheetNames;
        this.response = builder.response;
        this.map = builder.map;
    }


    //generate excel and send file outputStream to client
    public <T> void genExcel() throws IOException {

        //UFT-8编码防止中文不被解析
        String fileName = URLEncoder.encode(this.outputFileName.trim() + ".xlsx", "UTF8");
        //java 10 及以上可以采用StandardCharsets.UTF_8

        /* 输出方式一：先生成本地文件，再返回文件流 */
        //略

        //输出方式二：直接通过response获取servletOutputStream流返回
        response.setCharacterEncoding("UTF-8");
        /* 设置xlsx格式响应头及文件名 */
        // MIME types(IANA media types) 链接如下：
        // https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.addHeader("Content-Disposition", "attachment; filename=" + fileName);

        ServletOutputStream fos = response.getOutputStream();        //get response outputStream

        ExcelWriter writer = EasyExcel.write(fos).withTemplate(is).build();
        //fill data to corresponding sheet
        for (int index = 0; index < this.sheetNos.size(); index++) {
            WriteSheet sheet = EasyExcel.writerSheet(this.sheetNos.get(index), this.sheetNames.get(index))
                    .build();
            writer.fill(this.data.get(index), sheet);
            //map填入
            writer.fill(map, sheet);
        }
        writer.finish();        //close IO
        /* flush and close outputStream */
        fos.flush();
        fos.close();
    }

    //builder
    public static class ExcelBuilder<E2> {
        private final List<E2> data;
        private final InputStream is;
        private final String outputFileName;
        private final List<Integer> sheetNos;
        private final List<String> sheetNames;
        private final HttpServletResponse response;
        private Map<String,Object> map;

        /**
         * make builder with template file path
         * @param data  data to fill into sheet
         * @param templatePath  relative template file path
         * @param outputFileName    the output filename
         * @param sheetNo   sheet number. start with zero(0)
         * @param sheetName sheet name correspondence to sheetNo
         * @param response  HttpServletResponse
         */
        public ExcelBuilder(E2 data, String templatePath, String outputFileName,
                            Integer sheetNo, String sheetName,
                            HttpServletResponse response) throws IOException {
            //通过流来操作，防止项目打包成jar时获取不到内部文件
            ClassPathResource classPathResource = new ClassPathResource(templatePath);
            InputStream is = classPathResource.getInputStream();

            this.data = new ArrayList<E2>() {{
                add(data);
            }};
            this.is = is;
            this.outputFileName = outputFileName;
            this.sheetNos = new ArrayList<Integer>() {{
                add(sheetNo);
            }};
            this.sheetNames = new ArrayList<String>() {{
                add(sheetName);
            }};
            this.response = response;
        }

        /**
         * make builder with java File Object —— templateFile
         * @param data  data to fill into sheet
         * @param templateFile  template File Object
         * @param outputFileName    the output filename
         * @param sheetNo   sheet number. start with zero(0)
         * @param sheetName sheet name correspondence to sheetNo
         * @param response  HttpServletResponse
         * @throws FileNotFoundException
         */
        public ExcelBuilder(E2 data, File templateFile, String outputFileName,
                            Integer sheetNo, String sheetName,
                            HttpServletResponse response) throws FileNotFoundException {
            if (!templateFile.getName().endsWith(".xlsx")) {
                throw new RuntimeException(templateFile.getName() + "不是.xlsx文件，请选择xlsx文件作为模板");
            }
            this.data = new ArrayList<E2>() {{
                add(data);
            }};
            this.is = new FileInputStream(templateFile);
            this.outputFileName = outputFileName;
            this.sheetNos = new ArrayList<Integer>() {{
                add(sheetNo);
            }};
            this.sheetNames = new ArrayList<String>() {{
                add(sheetName);
            }};
            this.response = response;
        }

        /**
         * make builder with inputStream of templateFile
         * @param data  data to fill into sheet
         * @param is  template File inputstream
         * @param outputFileName    the output filename
         * @param sheetNo   sheet number. start with zero(0)
         * @param sheetName sheet name correspondence to sheetNo
         * @param response  HttpServletResponse
         */
        public ExcelBuilder(E2 data, InputStream is, String outputFileName,
                            Integer sheetNo, String sheetName,
                            HttpServletResponse response) {
            this.data = new ArrayList<E2>() {{
                add(data);
            }};
            this.is = is;
            this.outputFileName = outputFileName;
            this.sheetNos = new ArrayList<Integer>() {{
                add(sheetNo);
            }};
            this.sheetNames = new ArrayList<String>() {{
                add(sheetName);
            }};
            this.response = response;
        }

        /**
         * make builder with inputStream of templateFile
         * @param data  a list of data to fill into sheets
         * @param is  template File inputStream
         * @param outputFileName    the output filename
         * @param sheetNos   a list of sheet number. start with zero(0)
         * @param sheetNames    a list of sheet name correspondence to sheetNos
         * @param response  HttpServletResponse
         */
        public ExcelBuilder(List<E2> data, InputStream is, String outputFileName,
                            List<Integer> sheetNos, List<String> sheetNames,
                            HttpServletResponse response) {
            this.data = data;
            this.is = is;
            this.outputFileName = outputFileName;
            this.sheetNos = sheetNos;
            this.sheetNames = sheetNames;
            this.response = response;
        }

        // add additional sheet
        public ExcelBuilder<E2> addSheet(Integer sheetNo, String sheetName,E2 data) {
            this.sheetNos.add(sheetNo);
            this.sheetNames.add(sheetName);
            this.data.add(data);
            return this;
        }

        public ExcelBuilder<E2> addKV(String k,Object v){
            if(this.map == null){
                this.map = new HashMap<>();
            }
            map.put(k,v);
            return this;
        }

        public ExcelBuilder<E2> setMap(Map<String,Object> map){
            this.map = map;
            return this;
        }

        public XlsxUtils<E2> build() {
            return new XlsxUtils<E2>(this);
        }
    }
}
