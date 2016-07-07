package com.cloudcode.doc.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import com.cloudcode.framework.utils.StringUtils;
import com.cloudcode.framework.utils.UUID;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

public class World2PDF {
	public static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
	public static final int wdFormatPDF = 17;// PDF 格式
	public static void main(String[] args) {
        createPDF();
        createPDF();
    }
 
    private static void createPDF() {
        try {
            long start = System.currentTimeMillis();
 
            // 1) Load DOCX into WordprocessingMLPackage
            InputStream is = new FileInputStream(new File(
                    "docx/1.docx"));
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(is);
            Mapper fontMapper = new IdentityPlusMapper();  
            fontMapper.getFontMappings().put("华文行楷", PhysicalFonts.getPhysicalFonts().get("STXingkai"));  
            fontMapper.getFontMappings().put("华文仿宋", PhysicalFonts.getPhysicalFonts().get("STFangsong"));  
            fontMapper.getFontMappings().put("隶书", PhysicalFonts.getPhysicalFonts().get("LiSu"));  
            fontMapper.getFontMappings().put("宋体", PhysicalFonts.getPhysicalFonts().get("SimSun"));  
            wordMLPackage.setFontMapper(fontMapper);  
            // 2) Prepare Pdf settings
            PdfSettings pdfSettings = new PdfSettings();
            Docx4jProperties.getProperties().setProperty("docx4j.Log4j.Configurator.disabled", "true");
            //Log4jConfigurator.configure();            
           // org.docx4j.convert.out.pdf.viaXSLFO.Conversion.log.setLevel(Level.OFF);
            // 3) Convert WordprocessingMLPackage to Pdf
            OutputStream out = new FileOutputStream(new File(
                    "pdf/HelloWorld.pdf"));
            PdfConversion converter = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(
                    wordMLPackage);
            converter.output(out, pdfSettings);
 
            System.err.println("Generate pdf/HelloWorld.pdf with "
                    + (System.currentTimeMillis() - start) + "ms");
 
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
    public static void wordToPdfWaterMark(String docpath, String pdfpath,
			String waterMarkName, String logoPath, boolean isdeletetempfile) {
		try {
			int index = docpath.lastIndexOf("/");
			String uuid = UUID.generateUUID();
			String ttt = "/temp"+uuid+"/";
			if (index == -1) {
				index = docpath.lastIndexOf("\\");
				ttt = "\\temp"+uuid+"\\";
			}
			String wenjianjia = docpath.substring(0, index) + ttt;

			String wenjianming = "";

			String temp = wenjianming + UUID.generateUUID();

			File wenjianjiafile = new File(wenjianjia);
			wenjianjiafile.mkdir();

			WordToPDFByJacob(docpath, wenjianjia + temp + ".pdf");
			PdfReader reader = new PdfReader(wenjianjia + temp + ".pdf");// 输入PDF
			PdfStamper stamp = new PdfStamper(reader, new FileOutputStream(
					pdfpath));// 输出PDF	
			if(StringUtils.isEmpty(logoPath)){
				PdfUtils.addWatermark(stamp, waterMarkName);
				stamp.close(); 
			}else{
				Rectangle rect = null;
				if (reader.getPageSize(1) != null)
					rect = new Rectangle(reader.getPageSize(1));
				else
					rect = new Rectangle(PageSize.A4);
				PdfUtils.addWatermark(stamp, rect, waterMarkName, logoPath);
			}
			if (isdeletetempfile == true) {
				File[] files = wenjianjiafile.listFiles();
				for (File file2 : files) {
					file2.delete();
				}
				wenjianjiafile.delete();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
    public static void wordToPdfWaterMark(String docpath, String pdfpath,
			String waterMarkName, boolean isdeletetempfile) {
    	wordToPdfWaterMark(docpath, pdfpath, waterMarkName, null, isdeletetempfile);
	}
	public static void WordToPDFByJacob(String sfileName, String toFileName) {	
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);

			Dispatch docs = app.getProperty("Documents").toDispatch();		
			Dispatch doc = Dispatch.call(docs,//
					"Open", //
					sfileName,// FileName
					false,// ConfirmConversions
					true // ReadOnly
					).toDispatch();
			Dispatch word=Dispatch.get(doc, "Revisions").toDispatch();
			Dispatch.call(word, "AcceptAll");
			File tofile = new File(toFileName);
			if (tofile.exists()) {
				tofile.delete();
			}
			Dispatch.call(doc,//
					"SaveAs", //
					toFileName, // FileName
					wdFormatPDF);

			Dispatch.call(doc, "Close", false);

			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");

		} catch (Exception e) {
			System.out.println("========Error:Word文档转换Pdf失败：" + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", wdDoNotSaveChanges);
		}

	}
}

