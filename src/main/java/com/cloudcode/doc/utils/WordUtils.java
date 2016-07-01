package com.cloudcode.doc.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.jsoup.Jsoup;
import org.jsoup.helper.StringUtil;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.cloudcode.doc.utils.inter.IHtml;
import com.cloudcode.framework.utils.IOUtils;
import com.cloudcode.framework.utils.PropertiesUtil;
import com.cloudcode.framework.utils.UUID;
import com.google.common.base.Charsets;
import com.google.common.io.Files;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WordUtils implements IHtml {
	private static final Logger logger = LoggerFactory
			.getLogger(WordUtils.class);

	public String getContent(String filePath) {

		return null;
	}

	public static final int HTML_WORD = 1;
	public static final int WORD_HTML = 8;
	public static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
	public static final int wdFormatPDF = 17;// PDF 格式
	public static  PropertiesUtil propertiesUtil = new PropertiesUtil("");
	public static boolean htmlToWord(String htmlfile, String docfile) {
		boolean b = false;
		ActiveXComponent app = new ActiveXComponent("Word.Application");
		try {
			app.setProperty("Visible", new Variant(false));
			Dispatch docs = app.getProperty("Documents").toDispatch();
			Dispatch doc = Dispatch.invoke(
					docs,
					"Open",
					1,
					new Object[] { htmlfile, new Variant(false),
							new Variant(true) }, new int[1]).toDispatch();
			Dispatch.invoke(doc, "SaveAs", 1, new Object[] { docfile,
					new Variant(1) }, new int[1]);
			Variant f = new Variant(false);
			Dispatch.call(doc, "Close", f);
			b = true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			app.invoke("Quit", new Variant[0]);
		}
		return b;
	}

	public static boolean wordToHtml(String docfile, String htmlfile) {
		boolean b = false;
		ActiveXComponent app = new ActiveXComponent("Word.Application");
		try {
			app.setProperty("Visible", new Variant(false));
			Dispatch docs = app.getProperty("Documents").toDispatch();
			// Dispatch.call(docs, "wdPageBreak");
			Dispatch doc = Dispatch.invoke(
					docs,
					"Open",
					1,
					new Object[] { docfile, new Variant(false),
							new Variant(true) }, new int[1]).toDispatch();
			// Selection.InsertBreak Type:=wdPageBreak
			// Dispatch selection = Dispatch.get(objWord,
			// "Selection").toDispatch();

			Dispatch.invoke(doc, "SaveAs", 1, new Object[] { htmlfile,
					new Variant(8) }, new int[1]);
			Variant f = new Variant(false);
			Dispatch.call(doc, "Close", f);
			b = true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			app.invoke("Quit", new Variant[0]);
		}
		return b;
	}

	public static String getHtmlContentByWord(String docfile, String htmlfile,
			String systemPath) {
		String html = "";
		try {
			wordToHtml(docfile, htmlfile);
			html = getHtmlInnerContent(htmlfile, systemPath);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return html;
	}

	public static String getHtmlInnerContent(String htmlfile, String systemPath)
			throws UnsupportedEncodingException, IOException {
		String docpath = getTempPath();

		String html = "";
		html = Files.toString(new File(htmlfile), Charsets.ISO_8859_1);
		html = new String(html.getBytes("iso-8859-1"), "gb2312");

		Document doc = Jsoup.parse(html, "UTF-8");
		Elements elements = doc.select("body").first().children();
		String content = "";
		String s = elements.toString();
		for (Element el : elements) {
			Elements imgs = el.getElementsByTag("img");
			for (Element img : imgs) {
				img.attr("src", systemPath + img.attr("src"));
			}
			content = content + replaceHtml(el.toString());
		}
		File htmlFile = new File(htmlfile);
		if (htmlFile.exists()) {
			htmlFile.deleteOnExit();
		}
		return content.toString();
	}

	public static String replaceHtml(String html) {
		html = html.replaceAll("&lt;", "<").replaceAll("&gt;", ">")
				.replaceAll(" ", "\n");
		Pattern patterns = Pattern.compile("(?s)<!--.*?-->");
		Matcher ms = patterns.matcher(html);
		html = ms.replaceAll("");
		html = html.replaceAll("<span[\\s]*?style[^>]*?>&nbsp;</span>", "<br>");
		return html;
	}

	public static String getTempPath() {
		String path = WordUtils.class.getClassLoader().getResource("/")
				.getPath().substring(1);
		int index = path.lastIndexOf("/WEB-INF/classes");
		if (index > -1) {
			path = path.substring(0, index);
		}
		path = path + "/temp" + "/";
		return path;
	}

	// public static void main(String[] args) {
	// wordToHtml("\\Users\\lijian\\Documents\\测试.docx",
	// "\\Users\\lijian\\Documents\\result11.html");
	// }
	public static String readDoc(String doc) throws Exception {
		FileInputStream in = new FileInputStream(doc);
		org.textmining.text.extraction.WordExtractor extractor = null;
		String text = null;
		extractor = new org.textmining.text.extraction.WordExtractor();
		text = extractor.extractText(in);
		return text;
	}

	public static void read(String docFilePath) {
		// 建立ActiveX部件
		ActiveXComponent wordCom = new ActiveXComponent("Word.Application");
		// word应用程序不可见
		wordCom.setProperty("Visible", false);
		// 禁用宏
		wordCom.setProperty("AutomationSecurity", new Variant(3));
		try {
			// 返回wrdCom.Documents的Dispatch
			Dispatch wrdDocs = wordCom.getProperty("Documents").toDispatch();// Documents表示word的所有文档窗口（word是多文档应用程序）
			// 调用wrdCom.Documents.Open方法打开指定的word文档，返回wordDoc
			String password = "123";
			Dispatch wordDoc = Dispatch.call(wrdDocs, "Open", docFilePath,
					false,// ConfirmConversions
					true, false, new Variant(password)).toDispatch();
			Dispatch selection = Dispatch.get(wordCom, "Selection")
					.toDispatch();
			setHeaderContent(wrdDocs, "$$");
			// Dispatch.call(selection, "InsertBreak" , new Variant(7) );
			int pages = Integer.parseInt(Dispatch.call(selection,
					"information", 4).toString());// 总页数 //显示修订内容的最终状态
			System.out.println(pages);
			for (int i = 1; i <= pages; i++) {
				// System.out.println(getParagraphs(wordDoc, i));
			}
			// Dispatch.call(wordDoc, "AcceptAllRevisionsShown");
			// processId = processManager.findPid(PROCESS_COMMANDLINE);
			// return true;
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	public static String getParagraphs(Dispatch doc, int paragraphsIndex) {
		String ret = "";
		Dispatch paragraphs = Dispatch.get(doc, "Paragraphs").toDispatch(); // 所有段落
		int paragraphCount = Dispatch.get(paragraphs, "Count").getInt(); // 一共的段落数
		Dispatch paragraph = null;
		Dispatch range = null;
		if (paragraphCount > paragraphsIndex && 0 < paragraphsIndex) {
			paragraph = Dispatch.call(paragraphs, "Item",
					new Variant(paragraphsIndex)).toDispatch();
			range = Dispatch.get(paragraph, "Range").toDispatch();
			ret = Dispatch.get(range, "Text").toString();
		}
		return ret;
	}

	public static void setHeaderContent(Dispatch doc, String cont) {
		Dispatch activeWindow = Dispatch.get(doc, "ActiveWindow").toDispatch();
		Dispatch view = Dispatch.get(activeWindow, "View").toDispatch();
		// Dispatch seekView = Dispatch.get(view, "SeekView").toDispatch();
		Dispatch.put(view, "SeekView", new Variant(9)); // wdSeekCurrentPageHeader-9

		Dispatch headerFooter = Dispatch.get(doc, "HeaderFooter").toDispatch();
		Dispatch range = Dispatch.get(headerFooter, "Range").toDispatch();
		Dispatch.put(range, "Text", new Variant(cont));
		// String content = Dispatch.get(range, "Text").toString();
		Dispatch font = Dispatch.get(range, "Font").toDispatch();

		Dispatch.put(font, "Name", new Variant("楷体_GB2312"));
		Dispatch.put(font, "Bold", new Variant(true));
		// Dispatch.put(font, "Italic", new Variant(true));
		// Dispatch.put(font, "Underline", new Variant(true));
		Dispatch.put(font, "Size", 9);

		Dispatch.put(view, "SeekView", new Variant(0)); // wdSeekMainDocument-0恢复视图;
	}

	public static void main(String[] args) {
		WordUtils.read("c://test//test.doc");
	}

	public static void main2(String[] args) {
		try {
			// String text =
			// WordToHtml.readDoc("/Users/lijian/Documents/e.doc");
			/*
			 * HWPFDocumentCore wordDocument = WordToHtmlUtils .loadDoc(new
			 * FileInputStream("c://test//test.doc"));
			 * 
			 * WordToHtmlConverter wordToHtmlConverter = new
			 * WordToHtmlConverter(
			 * DocumentBuilderFactory.newInstance().newDocumentBuilder()
			 * .newDocument());
			 * wordToHtmlConverter.processDocument(wordDocument);
			 * org.w3c.dom.Document htmlDocument = wordToHtmlConverter
			 * .getDocument(); ByteArrayOutputStream out = new
			 * ByteArrayOutputStream(); DOMSource domSource = new
			 * DOMSource(htmlDocument); StreamResult streamResult = new
			 * StreamResult(out);
			 * 
			 * TransformerFactory tf = TransformerFactory.newInstance();
			 * Transformer serializer = tf.newTransformer();
			 * serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			 * serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			 * serializer.setOutputProperty(OutputKeys.METHOD, "html");
			 * serializer.transform(domSource, streamResult); out.close();
			 * 
			 * String result = new String(out.toByteArray());
			 * System.out.println(result);
			 */
			// System.out.println(text);
			Runtime.getRuntime().exec("taskKill /F /IM winword.exe");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/***************************************************************************
	 * 删除书签
	 * 
	 * @param mark
	 *            书签名
	 * @param info
	 *            可替换
	 * @return
	 */
	public boolean deleteBookMark(String markKey, String info) throws Exception {
		ActiveXComponent word = new ActiveXComponent("Word.Application");
		word.setProperty("Visible", new Variant(false));
		Dispatch activeDocument = word.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch bookMarks = word.call(activeDocument, "Bookmarks")
				.toDispatch();
		boolean isExists = word.call(bookMarks, "Exists", markKey).toBoolean();
		if (isExists) {

			Dispatch n = Dispatch.call(bookMarks, "Item", markKey).toDispatch();
			Dispatch.call(n, "Delete");

			return true;
		}
		return false;
	}

	/***************************************************************************
	 * 根据书签插入数据
	 * 
	 * @param bookMarkKey
	 *            书签名
	 * @param info
	 *            插入的数据
	 * @return
	 */

	public boolean intoValueBookMark(String bookMarkKey, String info)
			throws Exception {
		ActiveXComponent word = new ActiveXComponent("Word.Application");
		word.setProperty("Visible", new Variant(false));
		Dispatch activeDocument = word.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch bookMarks = word.call(activeDocument, "Bookmarks")
				.toDispatch();
		boolean bookMarkExist = word.call(bookMarks, "Exists", bookMarkKey)
				.toBoolean();
		if (bookMarkExist) {

			Dispatch rangeItem = Dispatch.call(bookMarks, "Item", bookMarkKey)
					.toDispatch();
			Dispatch range = Dispatch.call(rangeItem, "Range").toDispatch();
			Dispatch.put(range, "Text", new Variant(info));
			return true;
		}
		return false;
	}

	public static Map<String, String> replaceAllWordBookMark2(String docfile,
			Map<String, Object> map) {
		Map<String, String> bookMark = new HashMap<String, String>();
		ActiveXComponent word = new ActiveXComponent("Word.Application");
		try {
			word.setProperty("Visible", new Variant(true));
		} catch (Exception e) {
			e.printStackTrace();
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(true));
		}

		word.setProperty("AutomationSecurity", new Variant(3));
		Dispatch documents = word.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(documents, "Open", docfile).toDispatch();
		try {
			String value = "";
			Dispatch bookMarks = word.call(doc, "Bookmarks").toDispatch();
			int bCount = Dispatch.get(bookMarks, "Count").getInt();
			for (int i = 1; i <= bCount; i++) {
				try {
					value = "";
					Dispatch item = Dispatch.call(bookMarks, "Item", i)
							.toDispatch();
					String name = String.valueOf(Dispatch.get(item, "Name")
							.getString());// 读取书签命名
					Dispatch range = Dispatch.get(item, "Range").toDispatch();
					if (null != map && map.containsKey(name)) {
						if (null != map.get(name))
							value = map.get(name).toString();
					}
					Dispatch.put(range, "Text", new Variant(value));
					Dispatch.call(bookMarks, "Add", name, range);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			Dispatch.call(doc, "SaveAs");
			Dispatch.call(doc, "Close", new Variant(true));
			if (word != null)
				word.invoke("Quit", new Variant[] {});
		}
		return bookMark;
	}

	private static final String prefixStr = "${";
	private static final String suffixStr = "!''}";

	public static Map<String, String> replaceAllWordBookMarkToFtl(
			String docfile, Map<String, Object> map, boolean removeBookMark) {
		Map<String, String> bookMark = new HashMap<String, String>();
		ActiveXComponent word = new ActiveXComponent("Word.Application");
		try {
			word.setProperty("Visible", new Variant(true));
		} catch (Exception e) {
			e.printStackTrace();
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(true));
		}

		word.setProperty("AutomationSecurity", new Variant(3));
		Dispatch documents = word.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(documents, "Open", docfile).toDispatch();
		try {
			String value = "";
			Dispatch bookMarks = word.call(doc, "Bookmarks").toDispatch();
			int bCount = Dispatch.get(bookMarks, "Count").getInt();
			List<String> markLists = new ArrayList<String>();
			for (int i = 1; i <= bCount; i++) {
				try {

					Dispatch item = Dispatch.call(bookMarks, "Item", i)
							.toDispatch();
					String name = String.valueOf(Dispatch.get(item, "Name")
							.getString());
					value = prefixStr + name + suffixStr;
					Dispatch range = Dispatch.get(item, "Range").toDispatch();
					if (null != map && map.containsKey(name)) {
						if (null != map.get(name))
							value = map.get(name).toString();
					}
					Dispatch.put(range, "Text", new Variant(value));
					Dispatch.call(bookMarks, "Add", name, range);
					markLists.add(name);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			if (removeBookMark) {
				for (String key : markLists) {
					boolean exist = word.call(bookMarks, "Exists", key)
							.toBoolean();
					if (exist) {
						Dispatch removeDis = Dispatch.call(bookMarks, "Item",
								key).toDispatch();
						Dispatch.call(removeDis, "Delete");
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			Dispatch.call(doc, "SaveAs");
			Dispatch.call(doc, "Close", new Variant(true));
			if (word != null)
				word.invoke("Quit", new Variant[] {});
		}
		return bookMark;
	}

	public static void wordToPDF(String sfileName, String toFileName) {
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application");
			Dispatch docs = app.getProperty("Documents").toDispatch();
			Dispatch doc = Dispatch.call(docs,//
					"Open", //
					sfileName,// FileName
					false,// ConfirmConversions
					true // ReadOnly
					).toDispatch();

			Dispatch word = Dispatch.get(doc, "Revisions").toDispatch();
			Dispatch.call(word, "AcceptAll");

			Dispatch ActiveWindow = Dispatch.get(doc, "ActiveWindow")
					.toDispatch();
			Dispatch View = Dispatch.get(ActiveWindow, "View").toDispatch();
			Dispatch.put(View, "ShowRevisionsAndComments", false);
			Dispatch.put(View, "RevisionsView", 0);
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
			logger.error("微软wordToPDF 转换异常，转换文件为：" + sfileName + ",错误信息：：："
					+ e.getMessage());
			e.printStackTrace();
			System.out.println(" Error:微软wordToPDF 文档转换失败：" + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", wdDoNotSaveChanges);
		}

	}
	/**
	 * 修订文档
	 * deleteWordRevisions:(这里用一句话描述这个方法的作用). <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 * TODO(这里描述这个方法的执行流程 – 可选).<br/>
	 * TODO(这里描述这个方法的使用方法 – 可选).<br/>
	 * TODO(这里描述这个方法的注意事项 – 可选).<br/>
	 *
	 * @author cloudscode   ljzhuanjiao@Gmail.com
	 * @param wordName
	 * @since JDK 1.6
	 */
	public static void deleteWordRevisions(String wordName) {
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);

			Dispatch docs = app.getProperty("Documents").toDispatch();
			// System.out.println("打开文档..." + wordName);
			Dispatch doc = Dispatch.call(docs,//
					"Open", //
					wordName,// FileName
					false,// ConfirmConversions
					false // ReadOnly
					).toDispatch();
			Dispatch word = Dispatch.get(doc, "Revisions").toDispatch();
			Dispatch.call(word, "AcceptAll");
			Dispatch.call(doc, "SaveAs");
			Dispatch.call(doc, "Close", true);
		} catch (Exception e) {
			logger.error("deleteWordRevisions："  + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", wdDoNotSaveChanges);
		}
	}

	public static String flagUnderlineReplaceToTextarea(String oldStr) {
		oldStr = oldStr.replaceAll("(\r\n|\r|\n|\n\r)", " ");
		Pattern p = Pattern.compile("<a\\s.*?name=[^>]*>(.*?)</a>");
		String regEx_html="<[^>]+>";
		
		Pattern p3 = Pattern.compile("[_]{2,}");
		
		Pattern p2 = Pattern.compile(regEx_html,Pattern.CASE_INSENSITIVE);
		Matcher m = p.matcher(oldStr);
		int i = 0;
		while (m.find()) {
			String oldChar = m.group();
			//System.out.println("========oldChar====" + oldChar + ""); 
			int beginIndexName = m.group().indexOf("name=");
			int endIndexName = m.group().indexOf(">");
			if(beginIndexName>-1&&endIndexName>-1&&beginIndexName+5<endIndexName){
				String nameStr = m.group().substring(beginIndexName+5,endIndexName).trim();
				//System.out.println("========nameStr====[" + nameStr + "]"); 
				boolean isCode = WordUtils.isCode(nameStr);
				if (isCode){
					Matcher newStr=p2.matcher(oldChar);
					String value=newStr.replaceAll("");
					if(value==null){
                        value="";
					}
					String checkValue=value.replaceAll("(_|\\s*)", "");
					if("".equals(checkValue)){
						value="";
					 }
					i++;
					if(nameStr.indexOf("lawyee_")>-1){
						String newChar="<textarea spanId=span_"+i+" id=textarea_"+i+" onblur=\"var obj = document.getElementById('textarea_"+i+"\');var content = obj.value;if(content.length>0){content=content.replace(/^[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+|[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+$/g,'');} if(content.length>0){obj.style.display=\'none\';var no = obj.id.substring(9); var span = document.getElementById(\'span_\'+no);span.innerHTML =\'<font color=blue><U>\'+content+\'</U></font>\';span.style.display=\'\';}\" name="+nameStr+"  style=\"BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; MARGIN-TOP: 0px; WIDTH: auto; HEIGHT: auto; MARGIN-LEFT: 0px; OVERFLOW: visible; BORDER-TOP: black 1px solid; BORDER-RIGHT: black 1px solid; nowrap: \">"+value+"</textarea><span tabindex=\"0\" hidefocus=\"true\" style=\"none\" name=\"span_x\" id=span_"+i+" onfocus=\"var obj = document.getElementById(\'span_"+i+"\');var text = \'\';if (window.ActiveXObject){text = obj.outerText;} else if (window.XMLHttpRequest){text = obj.textContent;}if(text.length>0){obj.style.display=\'none\';var ta = document.getElementById(\'textarea_"+i+"\');ta.innerHTML =text;ta.style.display=\'\';ta.focus();}\"></span>";
						oldStr = oldStr.replace(oldChar, newChar);
					}else{
						String newChar="<textarea spanId=span_"+i+" id=textarea_"+i+" onblur=\"var obj = document.getElementById(\'textarea_"+i+"\');var content = obj.value;if(content.length>0){content=content.replace(/^[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+|[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+$/g,'');} if(content.length>0){obj.style.display=\'none\';var no = obj.id.substring(9); var span = document.getElementById(\'span_\'+no);span.innerHTML =\'<font color=blue><U>\'+content+\'</U></font>\';span.style.display=\'\';}\" name=\""+nameStr+"\"  style=\"BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; MARGIN-TOP: 0px; WIDTH: auto; HEIGHT: auto; MARGIN-LEFT: 0px; OVERFLOW: visible; BORDER-TOP: black 1px solid; BORDER-RIGHT: black 1px solid; nowrap: \">"+value+"</textarea><span tabindex=\"0\" hidefocus=\"true\" style=\"none\" name=\"span_x\" id=span_"+i+" onfocus=\"var obj = document.getElementById(\'span_"+i+"\');var text = \'\';if (window.ActiveXObject){text = obj.outerText;} else if (window.XMLHttpRequest){text = obj.textContent;}if(text.length>0){obj.style.display=\'none\';var ta = document.getElementById(\'textarea_"+i+"\');ta.innerHTML =text;ta.style.display=\'\';ta.focus();}\"></span>";
						oldStr = oldStr.replace(oldChar, newChar);
					}
				}
			}
		}
		return oldStr;
	}

	
	
	
	public static String flagUnderlineReplaceToTextarea(String oldStr,Map<String,String> bookMarkMap) {
		   if(bookMarkMap!=null){
			  Set<String> namelist=bookMarkMap.keySet();
			  int i=0;
			  for(String name:namelist){
				  String oldChar=WordUtils.getLawyeeHtmlBookMark(name);
				  String value=bookMarkMap.get(name);
				  if(value==null){
					  value="";
				  }
				  String checkValue=value.replaceAll("(_|\\s*)", "");
				  if("".equals(checkValue)){
						value="";
				   }
				  i++;
				  String newChar="<textarea spanId=span_"+i+" id=textarea_"+i+" onblur=\"var obj = document.getElementById(\'textarea_"+i+"\');var content = obj.value;if(content.length>0){content=content.replace(/^[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+|[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+$/g,'');} if(content.length>0){obj.style.display=\'none\';var no = obj.id.substring(9); var span = document.getElementById(\'span_\'+no);span.innerHTML =\'<font color=blue><U>\'+content+\'</U></font>\';span.style.display=\'\';}\" name=\""+name+"\"  style=\"BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; MARGIN-TOP: 0px; WIDTH: auto; HEIGHT: auto; MARGIN-LEFT: 0px; OVERFLOW: visible; BORDER-TOP: black 1px solid; BORDER-RIGHT: black 1px solid; nowrap: \">"+value+"</textarea><span tabindex=\"0\" hidefocus=\"true\" style=\"none\" name=\"span_x\" id=span_"+i+" onfocus=\"var obj = document.getElementById(\'span_"+i+"\');var text = \'\';if (window.ActiveXObject){text = obj.outerText;} else if (window.XMLHttpRequest){text = obj.textContent;}if(text.length>0){obj.style.display=\'none\';var ta = document.getElementById(\'textarea_"+i+"\');ta.innerHTML =text;ta.style.display=\'\';ta.focus();}\"></span>";
				  oldStr = oldStr.replace(oldChar, newChar);
			  }
		   }
	       return oldStr;
	}
	
	
	public static String flagUnderlineReplaceToTextarea(String oldStr,Map<String,String> bookMarkMap,Map<String,Object> bookMarkMapValue) {
		   if(bookMarkMap!=null){
			  Set<String> namelist=bookMarkMap.keySet();
			  int i=0;
			  for(String name:namelist){
				  String oldChar=WordUtils.getLawyeeHtmlBookMark(name);
				  String value=(String)bookMarkMapValue.get(name);
				  if(StringUtil.isBlank(value)){
					  value=bookMarkMap.get(name);
				  }
				  if(StringUtil.isBlank(value)){
					  value="";
				  }
				  String checkValue=value.replaceAll("(_|\\s*)", "");
				  if("".equals(checkValue)){
						value="";
				   }
				  i++;
				  String newChar="<textarea spanId=span_"+i+" id=textarea_"+i+" onblur=\"var obj = document.getElementById(\'textarea_"+i+"\');var content = obj.value;if(content.length>0){content=content.replace(/^[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+|[\\x09\\x0a\\x0b\\x0c\\x0d\\x20\\xa0\\u1680\\u180e\\u2000\\u2001\\u2002\\u2003\\u2004\\u2005\\u2006\\u2007\\u2008\\u2009\\u200a\\u2028\\u2029\\u202f\\u205f\\u3000]+$/g,'');} if(content.length>0){obj.style.display=\'none\';var no = obj.id.substring(9); var span = document.getElementById(\'span_\'+no);span.innerHTML =\'<font color=blue><U>\'+content+\'</U></font>\';span.style.display=\'\';}\" name=\""+name+"\"  style=\"BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; MARGIN-TOP: 0px; WIDTH: auto; HEIGHT: auto; MARGIN-LEFT: 0px; OVERFLOW: visible; BORDER-TOP: black 1px solid; BORDER-RIGHT: black 1px solid; nowrap: \">"+value+"</textarea><span tabindex=\"0\" hidefocus=\"true\" style=\"none\" name=\"span_x\" id=span_"+i+" onfocus=\"var obj = document.getElementById(\'span_"+i+"\');var text = \'\';if (window.ActiveXObject){text = obj.outerText;} else if (window.XMLHttpRequest){text = obj.textContent;}if(text.length>0){obj.style.display=\'none\';var ta = document.getElementById(\'textarea_"+i+"\');ta.innerHTML =text;ta.style.display=\'\';ta.focus();}\"></span>";
				  oldStr = oldStr.replace(oldChar, newChar);
			  }
		   }
	       return oldStr;
	}
	
	public static boolean isCode(String code){
		code = code.trim();
		Pattern p4 = Pattern.compile("^[a-zA-Z]+$");
		Matcher m4 = p4.matcher(code);
		if(m4.find()){
			return true;
		} else {
			if(code.indexOf("lawyee_")>-1){
				return true;
			}else{
				return false;
			}
		}
	}

	
	public static  String readHtmlToTxt(String html){
		String content ="";
		try {
			html="<html><head></head><body>"+html+"</body></html/>";
			byte[] htmlbyte = html.getBytes("gb2312");
			InputStream inputStream = new ByteArrayInputStream(htmlbyte);
			String htmlName=UUID.generateUUID();
			String docpath = getBasePath();
			String fullname = docpath + "\\" + htmlName+".html";
			File file=new File(fullname);
			OutputStream outputStream=new FileOutputStream(file);
			IOUtils.copyAndCloseIOStream(outputStream, inputStream);
			InputStream input=new FileInputStream(file);
			input.close();
			InputStream in=WordUtils.getHtmlToDocInputStream(htmlName, htmlName);
			content=WordUtils.readInputStream2String(in);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return content;
	}
	
	public static String getWordType(InputStream inputStream){
		try {
			if(POIFSFileSystem.hasPOIFSHeader(inputStream)){
				return "doc";
			}else if(POIXMLDocument.hasOOXMLHeader(inputStream)){
				return "docx";
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return "doc";
	}

	public static String readInputStream2String(InputStream inputStream){
		String result = "";
		String docpath = getBasePath();
		String id=UUID.generateUUID();
		String path=docpath+"\\" + id+".doc";
		File dirFile = new File(docpath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		try {
			if(POIFSFileSystem.hasPOIFSHeader(inputStream)){
				inputStream=WordUtils.acceptAllWordAllRevisions(inputStream,path);
				result=WordUtils.readWord2003(inputStream);
				File file=new File(path);
				if(file.exists()){
					file.delete();
				}
			}else if(POIXMLDocument.hasOOXMLHeader(inputStream)){
				inputStream=WordUtils.acceptAllWordAllRevisions(inputStream,path);
				result=WordUtils.readWord2007(inputStream);
				File file=new File(path);
				if(file.exists()){
					file.delete();
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		return result;
	}
	
   
	public static InputStream  acceptAllWordAllRevisions(InputStream inputStream,String path){
		String docpath = getBasePath();
		File dirFile = new File(docpath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		File file=new File(path);
		try {
			OutputStream out = new FileOutputStream(file);
			IOUtils.copyAndCloseIOStream(out,inputStream);
			WordUtils.deleteWordRevisions(path);
			InputStream in=new FileInputStream(file);
			return in;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return inputStream;
		}
	}

	
	public static String readWord2003(InputStream inputStream){
		String result = "";
		try {
			WordExtractor wordExtractor = new WordExtractor(inputStream);
			result = wordExtractor.getText();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	public static String readWord2007(InputStream inputStream){
		String result = "";
		try {
			OPCPackage opcp=OPCPackage.open(inputStream);
			POIXMLTextExtractor ex =new XWPFWordExtractor(opcp);
			result = ex.getText();
            opcp.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	
	
	public static FileInputStream getHtmlFromDocInputStream(String fileName,String hmtlName) {
		String docpath = getBasePath();
		String fullname = docpath + "\\" + fileName;
		String htmlname = docpath + "\\" + hmtlName + ".html";
		boolean flag = WordUtils.wordToHtml(fullname, htmlname);
		FileInputStream fis = null;
		if (flag) {
			try {
				fis = new FileInputStream(htmlname);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return fis;
	}
	
	
	public static Object[] getWordHtmlInputStream(String fileName,String hmtlName) {
		String docpath = getBasePath();
		String fullname = docpath + "\\" + fileName;
		String htmlname = docpath + "\\" + hmtlName + ".html";
		Map<String,String> bookMarkMap=WordUtils.restWordBookMark(fullname);
		boolean flag = WordUtils.wordToHtml(fullname, htmlname);
		FileInputStream fis = null;
		if (flag) {
			try {
				fis = new FileInputStream(htmlname);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		Object[] objects=new Object[2];
		objects[0]=fis;
		objects[1]=bookMarkMap;
		return objects;
	}
	
	public static Map<String,String> restWordBookMark(String docfile){
		Map<String,String> bookMark = new HashMap<String,String>();
		List<String> namelist=new ArrayList<String>();
		ActiveXComponent word = new ActiveXComponent("word.Application");
		word.setProperty("Visible", new Variant(false));
		word.setProperty("AutomationSecurity", new Variant(3));
		Dispatch documents = word.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(documents, "Open",docfile).toDispatch();
		try {
			//书签集合
			Dispatch bookMarks = word.call(doc, "Bookmarks").toDispatch();   
			int bCount = Dispatch.get(bookMarks, "Count").getInt();
			for (int i = 1; i <= bCount; i++) {
		         Dispatch item = Dispatch.call(bookMarks, "Item", i).toDispatch();  
		         String name = String.valueOf(Dispatch.get(item, "Name").getString());//读取书签命名
		         Dispatch range = Dispatch.get(item, "Range").toDispatch();
		         String value = String.valueOf(Dispatch.get(range, "Text").getString()); //读取书签文本
		         if(name!=null&&!"".equals(name)){
		              bookMark.put(name,value);
		              namelist.add(name);
		         }
		    }
			
			for(String name:namelist){
				WordUtils.intoValueBookMark(word, name, WordUtils.getLawyeeHtmlBookMark(name),false);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", true);
			Dispatch.call(word, "Quit");
		}
		return bookMark;
	}
	
	public static String getLawyeeHtmlBookMark(String name){
		return "$[lawyee]{"+name+"}";
	}
	
	
	public static boolean intoValueBookMark(ActiveXComponent word,String bookMarkKey,String info){
		return intoValueBookMark(word,bookMarkKey,info,false);
	}
	
	public static boolean intoValueBookMark(ActiveXComponent word,String bookMarkKey,String info,boolean isKeepMark){
		word.setProperty("Visible", new Variant(false));
		Dispatch activeDocument = word.getProperty("activeDocument").toDispatch();
		Dispatch bookMarks = word.call(activeDocument,"Bookmarks").toDispatch();
		boolean bookMarkExist = word.call(bookMarks, "Exists", bookMarkKey).toBoolean();
		if(bookMarkExist){
			Dispatch rangeItem = Dispatch.call(bookMarks, "item", bookMarkKey).toDispatch();
			Dispatch range =Dispatch.call(rangeItem, "Range").toDispatch();
			Dispatch.put(range,"Text",new Variant(info));
			return true;
		}
		return false;
	}
	
	
	/**
	 * 替换所以书签内容，为空替换为“   /   ”（保留书签标记）
	 * @param docfile
	 * @param valueMap
	 * @return
	 */
	public static boolean replaceAllWordBookMark(String docfile,Map<String,String> valueMap){
		boolean isResult=true;
		ActiveXComponent word = new ActiveXComponent("word.Application");
		word.setProperty("Visible", new Variant(false));
		word.setProperty("AutomationSecurity", new Variant(3));
		Dispatch documents = word.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(documents, "Open",docfile).toDispatch();
		try {
			//书签集合
			Dispatch bookMarks = word.call(doc, "Bookmarks").toDispatch();   
			int bCount = Dispatch.get(bookMarks, "Count").getInt();
			for (int i = 1; i <= bCount; i++) {
		         Dispatch item = Dispatch.call(bookMarks, "Item", i).toDispatch();  
		         String name = String.valueOf(Dispatch.get(item, "Name").getString());//读取书签命名
		         Dispatch range = Dispatch.get(item, "Range").toDispatch();
		         if(name!=null&&!"".equals(name)){
		              String value=valueMap.get(name);
		              if(StringUtil.isBlank(value)){
		            	  value="   /   ";
		              }
		              Dispatch.put(range,"Text",new Variant(value));
		              Dispatch.call(bookMarks, "Add",name,range);
		         }
		    }
		} catch (Exception e) {
			e.printStackTrace();
			isResult=false;
		} finally {
			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", true);
			Dispatch.call(word, "Quit");
		}
		return isResult;
	}
	
	
	
	/**
	 * 替换指定书签内容（保留书签标记）
	 * @param docfile
	 * @param valueMap
	 * @return
	 */
	public static boolean replaceSpecificBookMark(String docfile,Map<String,String> bookMarkMap){
		boolean isResult=true;
		ActiveXComponent word = new ActiveXComponent("word.Application");
		word.setProperty("Visible", new Variant(false));
		word.setProperty("AutomationSecurity", new Variant(3));
		Dispatch documents = word.getProperty("Documents").toDispatch();
		Dispatch doc = Dispatch.call(documents, "Open",docfile).toDispatch();
		try {
			 if(bookMarkMap!=null){
				 for(String name:bookMarkMap.keySet()){
					  String value=bookMarkMap.get(name);
		              if(!StringUtil.isBlank(value)){
		      			    //书签集合
		      			    Dispatch bookMarks = word.call(doc, "Bookmarks").toDispatch();  
			            	boolean bookMarkExist = word.call(bookMarks, "Exists", name).toBoolean();
			          		if(bookMarkExist){
			          			Dispatch rangeItem = Dispatch.call(bookMarks, "item", name).toDispatch();
			          			Dispatch range =Dispatch.call(rangeItem, "Range").toDispatch();
			          			Dispatch.put(range,"Text",new Variant(value));
			          			Dispatch.call(bookMarks, "Add",name,range);
			          		}
		              }
				 }
			 }
		} catch (Exception e) {
			e.printStackTrace();
			isResult=false;
		} finally {
			Dispatch.call(doc, "Save");
			Dispatch.call(doc, "Close", true);
			Dispatch.call(word, "Quit");
		}
		return isResult;
	}
	
	public static FileInputStream getHtmlToDocInputStream(String hmtlName,String fileName) {
		String docpath = getBasePath();
		String docfile = docpath + "\\" + fileName+".doc";
		String htmlfile = docpath + "\\" + hmtlName+".html";
		boolean flag = WordUtils.htmlToWord(htmlfile,docfile);
		FileInputStream fis = null;
		if (flag) {
			try {
				fis = new FileInputStream(docfile);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return fis;
	}
	public static String htmlToWord2(String contents,String fileName){
		String html ="";
		String fullname ="";
		String docfile = "";
		ByteArrayInputStream bais = null;
		FileOutputStream fos = null;
		try {
			html="<html><head></head><body>"+contents+"</body></html/>";
									
			byte[] htmlbyte = html.getBytes("gb2312");
			InputStream inputStream = new ByteArrayInputStream(htmlbyte);
			String htmlName=UUID.generateUUID();
			String docpath =WordUtils.getBasePath();
			File dirFile = new File(docpath);
     		if (!dirFile.exists()) {
     			dirFile.mkdirs();
     		}
			//String docpath =new PropertiesUtil("system.properties").readProperty("DOC.UPLOADPATH");
			fullname = docpath + "\\" + htmlName+".html";
			
			File file=new File(fullname);
			OutputStream outputStream=new FileOutputStream(file);
			IOUtils.copyAndCloseIOStream(outputStream, inputStream);
			InputStream input=new FileInputStream(file);
			input.close();
			if(StringUtils.isNotEmpty(fileName))
				docfile = docpath + "\\" + fileName+"temp.doc";
			else
				docfile = docpath + "\\" + htmlName+".doc";
			boolean flag = WordUtils.htmlToWord(fullname,docfile);
			if (flag) {
				return docfile;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		finally{
			try {
				if(bais !=null)
					bais.close();
				if(fos !=null)
					fos.close();
			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
		return docfile;
	}
	public static String getBasePath(){
		String docpath = propertiesUtil.readProperty("DOC.UPLOADPATH");
		return docpath;
	}
	public static  byte[] readHtmlToWordByPOI(String html){
		byte[] contens = null;
		ByteArrayInputStream bais = null;
		FileOutputStream fos = null;
		FileInputStream fis =null;
		ByteArrayOutputStream out = null;
		try {
			html="<html><head></head><body>"+html+"</body></html/>";
			byte[] htmlbyte = html.getBytes("gb2312");
			bais = new ByteArrayInputStream(htmlbyte);
			POIFSFileSystem poifs = new POIFSFileSystem();
			DirectoryEntry directory = poifs.getRoot();
			DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);
			String docpath =getBasePath();
			String docName=UUID.generateUUID();
			String fullname = docpath + "\\" + docName+".doc";
			fos = new FileOutputStream(fullname);
			poifs.writeFilesystem(fos);
			fis = new FileInputStream(fullname);
			out=new ByteArrayOutputStream();
			IOUtils.copyAndCloseIOStream(out,fis);
			contens = out.toByteArray();
			fis.close();
			bais.close();
			fos.close();
			out.close();
			return contens;
		} catch (Exception e) {
			e.printStackTrace();
		}
		finally{
			try {
				if(bais !=null)
					bais.close();
				if(fis !=null)
					fis.close();
				if(fos !=null)
					fos.close();
				if(out !=null)
					out.close();
			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
		return contens;
	}
	public static String getWebRootTempPath(){
		String path =WordUtils.class.getClassLoader().getResource("/").getPath().substring(1);
	    int index=path.lastIndexOf("/WEB-INF/classes");
	    if(index>-1){
	    	path=path.substring(0, index);
	    }
		path = path+"/temp"+"/";
		return path;
	}
	public static String readByteArray2String(byte[] content){
		String result = "";
		String docpath =WordUtils.getTempPath();
		String id=UUID.generateUUID();
		String path=docpath+"\\" + id+".doc";
		File dirFile = new File(docpath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		int count=0;
		InputStream inputStream = new ByteArrayInputStream(content);
		try {
			if(POIFSFileSystem.hasPOIFSHeader(inputStream)){
				count=1;
				inputStream=WordUtils.acceptAllWordAllRevisions(inputStream,path);
				result=WordUtils.readWord2003(inputStream);
			}else if(POIXMLDocument.hasOOXMLHeader(inputStream)){
				count=1;
				inputStream=WordUtils.acceptAllWordAllRevisions(inputStream,path);
				result=WordUtils.readWord2007(inputStream);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		if(StringUtil.isBlank(result)){
			try{
				if(count>0){
				  inputStream = new ByteArrayInputStream(content);
				}
				count=1;
			    inputStream=WordUtils.acceptOtherWordAllRevisions(inputStream,path);
				result=WordUtils.readWord2007(inputStream);
			}catch (Exception e){
				e.printStackTrace();
			}
			
			if(StringUtil.isBlank(result)){
				if(count>0){
				  inputStream = new ByteArrayInputStream(content);
				}
				try {
					Document doc = Jsoup.parse(inputStream, "gb2312", "");
					result = doc.text();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		File file=new File(path);
		if(file.exists()){
			file.delete();
		}
		return result ;
	}
	public static InputStream  acceptOtherWordAllRevisions(InputStream inputStream,String path){
		String docpath =WordUtils.getTempPath();
		File dirFile = new File(docpath);
		if (!dirFile.exists()) {
			dirFile.mkdirs();
		}
		String oldpath=docpath+"\\" + UUID.generateUUID()+".doc";
		File file=new File(path);
		File oldfile=new File(oldpath);
		try {
			OutputStream out = new FileOutputStream(oldfile);
			IOUtils.copyAndCloseIOStream(out,inputStream);
			WordUtils.saveAsFinishedWord2007(oldpath, path);
			oldfile.deleteOnExit();
			InputStream in=new FileInputStream(file);
			return in;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return inputStream;
		}
	}
	public static void saveAsFinishedWord2007(String wordName,String newWordName) {
		ActiveXComponent app = new ActiveXComponent("Word.Application");
		Dispatch doc=null;
		try {
			//app = new ActiveXComponent("Word.Application");
			//app.setProperty("Visible", false);
			Dispatch docs = app.getProperty("Documents").toDispatch();
			doc = Dispatch.call(docs,"Open", wordName,false,false).toDispatch();
			
			
			Dispatch word=Dispatch.get(doc, "Revisions").toDispatch();
			Dispatch.call(word, "AcceptAll");

			Dispatch ActiveWindow=Dispatch.get(doc, "ActiveWindow").toDispatch();
			Dispatch View=Dispatch.get(ActiveWindow, "View").toDispatch();
			Dispatch.put(View,"ShowRevisionsAndComments",false);
			Dispatch.put(View,"RevisionsView",0);

			Dispatch wordBasic=app.getProperty("WordBasic").toDispatch();
			Dispatch.call(wordBasic, "RemoveHeader");
			Dispatch.call(wordBasic, "RemoveFooter");
			try{
			   Dispatch.call(wordBasic, "DeleteAllCommentsInDoc");
			} catch (Exception e) {
			}
			//另存为word2007
			Dispatch.call(doc,"SaveAs",newWordName,12);
		} catch (Exception e) {
			System.out.println(e.getMessage());
			e.printStackTrace();
		}finally {
			WordUtils.closeDoc(doc, false);
			WordUtils.quitOffice(app);
		}
   }
	public static void closeDoc(Dispatch doc, boolean isSave) {
		if (doc != null) {
			try {
				Dispatch.call(doc, "Close", isSave);
			} catch (Exception e) {

			}
		}
	}
	public static void quitOffice(ActiveXComponent app) {
		if (app != null) {
			try {
				Dispatch.call(app, "Quit", wdDoNotSaveChanges);
				// ComThread.Release();
			} catch (Exception e) {
			}
		}
	}
}
