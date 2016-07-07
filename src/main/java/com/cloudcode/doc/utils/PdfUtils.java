/**
 * Project Name:Cloudcode-Doc
 * File Name:PdfUtils.java
 * Package Name:com.cloudcode.doc.utils
 * Date:2016-7-7上午11:48:56
 * Copyright (c) 2016, chenzhou1025@126.com All Rights Reserved.
 *
 */

package com.cloudcode.doc.utils;

import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

public class PdfUtils {

	public static void main(String[] args) throws Exception {
		String pdfFilePath = "f:/itext-demo.pdf";
		PdfReader pdfReader = new PdfReader("f:/itext-demo.pdf");
		// Get the PdfStamper object
		PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileOutputStream(
				"f:/itext-demo22.pdf"));
		addWatermark(pdfStamper, "cloudcode");

		pdfStamper.close();
	}

	public static void addWatermark(PdfStamper pdfStamper, String waterMarkName) {
		PdfContentByte content = null;
		BaseFont base = null;
		Rectangle pageRect = null;
		PdfGState gs = new PdfGState();
		try {
			// 设置字体
			base = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H",
					BaseFont.NOT_EMBEDDED);
		} catch (DocumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		try {
			if (base == null || pdfStamper == null) {
				return;
			}
			// 设置透明度为0.4
			gs.setFillOpacity(0.4f);
			gs.setStrokeOpacity(0.4f);
			int toPage = pdfStamper.getReader().getNumberOfPages();
			for (int i = 1; i <= toPage; i++) {
				pageRect = pdfStamper.getReader().getPageSizeWithRotation(i);
				// 计算水印X,Y坐标
				float x = pageRect.getWidth() / 2;
				float y = pageRect.getHeight() / 2;
				// 获得PDF最顶层
				content = pdfStamper.getOverContent(i);
				content.saveState();
				// set Transparency
				content.setGState(gs);
				content.beginText();
				content.setColorFill(BaseColor.GRAY);
				content.setFontAndSize(base, 60);
				// 水印文字成45度角倾斜
				content.showTextAligned(Element.ALIGN_CENTER, waterMarkName, x,
						y, 45);
				content.endText();
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			content = null;
			base = null;
			pageRect = null;
		}
	}

	public static void addWatermark(PdfStamper stamper,
			Rectangle pageRectangle, String waterMarkNameCenter, String logoPath)
			throws Exception, IOException {
		PdfContentByte content;
		BaseFont base = null;
		try {
			base = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",
					BaseFont.EMBEDDED);
		} catch (DocumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		// 计算水印X,Y坐标
		float x = pageRectangle.getWidth() / 2;
		float y = pageRectangle.getHeight() / 2;
		int i = 1;
		while (true) {
			content = stamper.getOverContent(i++);// 获得PDF最顶层
			if (null == content) {
				break;
			}
			/*
			 * if (txm != null && txm.length() > 0) { content.saveState(); Image
			 * wm = Image.getInstance(txm); wm.scaleAbsolute(150, 35);
			 * wm.setAbsolutePosition( pageRectangle.getWidth() -150- 60,
			 * pageRectangle.getHeight() -35 -5); content.addImage(wm);
			 * content.restoreState(); }
			 */
			if (logoPath != null) {
				content.saveState();
				Image wm = Image.getInstance(logoPath);
				wm.setAbsolutePosition(
						(pageRectangle.getWidth() - wm.getWidth()) / 2,
						pageRectangle.getHeight() / 2);
				PdfGState gs = new PdfGState();
				gs.setFillOpacity(0.2f);// 设置透明度为0.2
				content.setGState(gs);
				content.addImage(wm);
				content.restoreState();
			}
			content.saveState();
		}

		stamper.close();
	}
}
