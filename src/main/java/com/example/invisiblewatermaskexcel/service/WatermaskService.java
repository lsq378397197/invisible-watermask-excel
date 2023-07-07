package com.example.invisiblewatermaskexcel.service;


import com.groupdocs.watermark.Watermarker;
import com.groupdocs.watermark.common.HorizontalAlignment;
import com.groupdocs.watermark.common.VerticalAlignment;
import com.groupdocs.watermark.contents.SpreadsheetContent;
import com.groupdocs.watermark.options.SpreadsheetBackgroundWatermarkOptions;
import com.groupdocs.watermark.options.SpreadsheetLoadOptions;
import com.groupdocs.watermark.options.SpreadsheetWatermarkModernWordArtOptions;
import com.groupdocs.watermark.options.SpreadsheetWatermarkShapeOptions;
import com.groupdocs.watermark.watermarks.Font;
import com.groupdocs.watermark.watermarks.ImageWatermark;
import com.groupdocs.watermark.watermarks.SizingType;
import com.groupdocs.watermark.watermarks.TextWatermark;
import org.springframework.stereotype.Service;

@Service
public class WatermaskService {
    public void add(String inputExcelPath, String outputExcelPath) {
        SpreadsheetLoadOptions loadOptions = new SpreadsheetLoadOptions();
        // Constants.InSpreadsheetXlsx is an absolute or relative path to your document. Ex: "C:\\Docs\\spreadsheet.xlsx"
        Watermarker watermarker = new Watermarker(inputExcelPath, loadOptions);

        TextWatermark textWatermark = new TextWatermark("Test watermark", new Font("Arial", 8));
        SpreadsheetWatermarkModernWordArtOptions options = new SpreadsheetWatermarkModernWordArtOptions();
        options.setWorksheetIndex(0);

        watermarker.add(textWatermark, options);
        watermarker.save(outputExcelPath);

        watermarker.close();
    }

    public void addV2(String inputExcelPath,String imagePath, String outputExcelPath){
        SpreadsheetLoadOptions loadOptions = new SpreadsheetLoadOptions();
        // Constants.InSpreadsheetXlsx is an absolute or relative path to your document. Ex: "C:\\Docs\\spreadsheet.xlsx"
        Watermarker watermarker = new Watermarker(inputExcelPath, loadOptions);

        // Add text watermark to the first worksheet
//        TextWatermark textWatermark = new TextWatermark("Test watermark", new Font("Arial", 8));
//        SpreadsheetWatermarkShapeOptions textWatermarkOptions = new SpreadsheetWatermarkShapeOptions();
//        textWatermarkOptions.setWorksheetIndex(0);
//        watermarker.add(textWatermark, textWatermarkOptions);

        // Add image watermark to the second worksheet
        ImageWatermark imageWatermark = new ImageWatermark(imagePath);

        SpreadsheetWatermarkShapeOptions imageWatermarkOptions = new SpreadsheetWatermarkShapeOptions();
        imageWatermarkOptions.setWorksheetIndex(1);
        watermarker.add(imageWatermark, imageWatermarkOptions);

        watermarker.save(outputExcelPath);

        watermarker.close();
        imageWatermark.close();
    }

    public void addV3(String inputExcelPath, String outputExcelPath){
        SpreadsheetLoadOptions loadOptions = new SpreadsheetLoadOptions();
// Constants.InSpreadsheetXlsx is an absolute or relative path to your document. Ex: "C:\\Docs\\spreadsheet.xlsx"
        Watermarker watermarker = new Watermarker(inputExcelPath, loadOptions);

        TextWatermark watermark = new TextWatermark("Test watermark", new Font("Segoe UI", 19));
        watermark.setHorizontalAlignment(HorizontalAlignment.Center);
        watermark.setVerticalAlignment(VerticalAlignment.Center);
        watermark.setRotateAngle(45);
        watermark.setSizingType(SizingType.ScaleToParentDimensions);
        watermark.setScaleFactor(0.5);
        watermark.setOpacity(0.1);

        SpreadsheetContent content = watermarker.getContent(SpreadsheetContent.class);
        SpreadsheetBackgroundWatermarkOptions options = new SpreadsheetBackgroundWatermarkOptions();
        options.setBackgroundWidth(content.getWorksheets().get_Item(0).getContentAreaWidthPx()); /* set background width */
        options.setBackgroundHeight(content.getWorksheets().get_Item(0).getContentAreaHeightPx()); /* set background height */
        watermarker.add(watermark, options);

        watermarker.save(outputExcelPath);

        watermarker.close();
    }

    public static void main(String[] args) {
        WatermaskService instance = new WatermaskService();
//        instance.add("E:/test.xlsx","E:/test-output.xlsx");
//        instance.addV2("E:/test.xlsx","E:/watermask.png","E:/test-output-v2.xlsx");
        instance.addV3("E:/test.xlsx","E:/test-output-v3.xlsx");
    }
}
