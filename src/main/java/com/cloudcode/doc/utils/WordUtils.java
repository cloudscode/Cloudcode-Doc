package com.cloudcode.doc.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.textmining.text.extraction.WordExtractor;

import com.cloudcode.doc.utils.inter.IHtml;
import com.cloudcode.framework.utils.FileUtils;
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
		WordExtractor extractor = null;
		String text = null;
		extractor = new WordExtractor();
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

	public static Map<String, String> replaceAllWordBookMark(String docfile,
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

			Dispatch word=Dispatch.get(doc, "Revisions").toDispatch();
			Dispatch.call(word, "AcceptAll");
			
			Dispatch ActiveWindow=Dispatch.get(doc, "ActiveWindow").toDispatch();
			Dispatch View=Dispatch.get(ActiveWindow, "View").toDispatch();
			Dispatch.put(View,"ShowRevisionsAndComments",false);
			Dispatch.put(View,"RevisionsView",0);			
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
			logger.error("微软wordToPDF 转换异常，转换文件为："+sfileName+",错误信息：：："+e.getMessage());
			e.printStackTrace();
			System.out.println(" Error:微软wordToPDF 文档转换失败：" + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", wdDoNotSaveChanges);
		}

	}
}
