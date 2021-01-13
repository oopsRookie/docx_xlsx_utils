package com.oopsRookie.util.excel;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * if you need add watermark, you have to set `inMemory(true)`
 * so far，I only find XSSFWorkbook does support watermark adding，
 * but other workbook like SXSSFWorkbook, I have no idea about it.
 *
 * ps: easyExcel's default type is SXSSFWorkbook(for efficiency),
 * so you should set `inMemory(true)` to change workbook type to XSSFWorkbook.
 */
@RequiredArgsConstructor
public class WaterMarkHandler implements SheetWriteHandler {

    private final String WATER_MARK;

    /**
     * add watermark and return byte output stream
     * @param content watermark name
     * @return ByteArrayOutputStream
     * @throws IOException
     */
    public static ByteArrayOutputStream createWaterMark(String content) throws IOException {
        int width = 200;
        int height = 150;
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        String fontType = "Microsoft Yahei";
        int fontStyle = Font.BOLD;
        int fontSize = 20;
        Font font = new Font(fontType, fontStyle, fontSize);
        Graphics2D g2d = image.createGraphics();
        image = g2d.getDeviceConfiguration().createCompatibleImage(width, height, Transparency.TRANSLUCENT);
        g2d.dispose();
        g2d = image.createGraphics();
        g2d.setColor(new Color(0, 0, 0, 20));
        g2d.setStroke(new BasicStroke(1));
        g2d.setFont(font);
        g2d.rotate(-0.5, (double) image.getWidth() / 2, (double) image.getHeight() / 2);
        FontRenderContext context = g2d.getFontRenderContext();
        Rectangle2D bounds = font.getStringBounds(content, context);
        double x = (width - bounds.getWidth()) / 2;
        double y = (height - bounds.getHeight()) / 2;
        double ascent = -bounds.getY();
        double baseY = y + ascent;
        g2d.drawString(content, (int) x, (int) baseY);
        g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
        g2d.dispose();
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        ImageIO.write(image, "png", os);
        return os;
    }


    /**
     * add a watermark on a sheet
     *
     * @param sheet excel sheet
     * @param bytes watermark name
     */
    public static void putWaterRemarkToExcel(XSSFSheet sheet, byte[] bytes) {
        XSSFWorkbook workbook = sheet.getWorkbook();
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        String rID = sheet.addRelation(null, XSSFRelation.IMAGES, workbook.getAllPictures().get(pictureIdx))
                .getRelationship().getId();
        sheet.getCTWorksheet().addNewPicture().setId(rID);
    }

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
    }

    @SneakyThrows
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        try (ByteArrayOutputStream waterMark = createWaterMark(WATER_MARK)){
            XSSFSheet sheet = (XSSFSheet) writeSheetHolder.getSheet();
            putWaterRemarkToExcel(sheet, waterMark.toByteArray());
        }
    }
}
