package com.cloudcode.doc.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.textmining.text.extraction.WordExtractor;

import com.cloudcode.doc.utils.inter.IHtml;
import com.google.common.base.Charsets;
import com.google.common.io.Files;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WordToHtml implements IHtml {

	public String getContent(String filePath) {

		return null;
	}
	public static final int HTML_WORD = 1;
	  public static final int WORD_HTML = 8;

	  public static boolean htmlToWord(String htmlfile, String docfile)
	  {
	    boolean b = false;
	    ActiveXComponent app = new ActiveXComponent("Word.Application");
	    try {
	      app.setProperty("Visible", new Variant(false));
	      Dispatch docs = app.getProperty("Documents").toDispatch();
	      Dispatch doc = Dispatch.invoke(docs, "Open", 1, new Object[] { htmlfile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
	      Dispatch.invoke(doc, "SaveAs", 1, new Object[] { docfile, new Variant(1) }, new int[1]);
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

	  public static boolean wordToHtml(String docfile, String htmlfile)
	  {
	    boolean b = false;
	    ActiveXComponent app = new ActiveXComponent("Word.Application");
	    try {
	      app.setProperty("Visible", new Variant(false));
	      Dispatch docs = app.getProperty("Documents").toDispatch();
	     // Dispatch.call(docs, "wdPageBreak");
	      Dispatch doc = Dispatch.invoke(docs, "Open", 1, new Object[] { docfile, new Variant(false), new Variant(true) }, new int[1]).toDispatch();
	     //  Selection.InsertBreak Type:=wdPageBreak
	     // Dispatch selection = Dispatch.get(objWord, "Selection").toDispatch();
	     
	      
	      Dispatch.invoke(doc, "SaveAs", 1, new Object[] { htmlfile, new Variant(8) }, new int[1]);
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

	  public static String getHtmlContentByWord(String docfile, String htmlfile, String systemPath) {
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
	    throws UnsupportedEncodingException, IOException
	  {
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

	  public static String replaceHtml(String html)
	  {
	    html = html.replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll(" ", "\n");
	    Pattern patterns = Pattern.compile("(?s)<!--.*?-->");
	    Matcher ms = patterns.matcher(html);
	    html = ms.replaceAll("");
	    html = html.replaceAll("<span[\\s]*?style[^>]*?>&nbsp;</span>", "<br>");
	    return html;
	  }
	  public static String getTempPath() {
		    String path = WordToHtml.class.getClassLoader().getResource("/").getPath().substring(1);
		    int index = path.lastIndexOf("/WEB-INF/classes");
		    if (index > -1) {
		      path = path.substring(0, index);
		    }
		    path = path + "/temp" + "/";
		    return path;
		  }
//	  public static void main(String[] args) {
//	    wordToHtml("\\Users\\lijian\\Documents\\测试.docx", "\\Users\\lijian\\Documents\\result11.html");
//	  }
	  public static String readDoc(String doc) throws Exception {  
	        FileInputStream in = new FileInputStream(doc);  
	        WordExtractor extractor = null;  
	        String text = null;  
	        extractor = new WordExtractor();  
	        text = extractor.extractText(in);  
	        return text;  
	    }  
	  
	  public static void read(String docFilePath){
		// 建立ActiveX部件 
		  ActiveXComponent  wordCom = new ActiveXComponent("Word.Application"); 
		  //word应用程序不可见 
		  wordCom.setProperty("Visible", false); 
		  // 禁用宏  
		  wordCom.setProperty("AutomationSecurity", new Variant(3)); 
		  try { 
		  // 返回wrdCom.Documents的Dispatch 
		  Dispatch wrdDocs = wordCom.getProperty("Documents").toDispatch();//Documents表示word的所有文档窗口（word是多文档应用程序） 
		  // 调用wrdCom.Documents.Open方法打开指定的word文档，返回wordDoc 
		  String password ="123"; 
		  Dispatch  wordDoc = Dispatch.call(wrdDocs, "Open", docFilePath, false,// ConfirmConversions  
		  true, false, new Variant(password)).toDispatch(); 
		  Dispatch selection = Dispatch.get(wordCom, "Selection").toDispatch(); 
		  setHeaderContent(wrdDocs, "$$");
		  //Dispatch.call(selection,  "InsertBreak" ,  new Variant(7) );
		  int pages = Integer.parseInt(Dispatch.call(selection,"information",4).toString());//总页数 //显示修订内容的最终状态
		  System.out.println(pages);
		  for(int i=1;i<=pages;i++){
			  //System.out.println(getParagraphs(wordDoc, i));
		  }
		  //Dispatch.call(wordDoc, "AcceptAllRevisionsShown"); 
		  //processId = processManager.findPid(PROCESS_COMMANDLINE); 
		 // return true; 
		  } catch (Exception ex) { 
			  ex.printStackTrace(); 		 
		  } 
	  }
	  public static String getParagraphs( Dispatch doc,int paragraphsIndex){  
	        String ret = "";  
	        Dispatch paragraphs = Dispatch.get(doc, "Paragraphs").toDispatch(); // 所有段落  
	        int paragraphCount = Dispatch.get(paragraphs, "Count").getInt();            // 一共的段落数  
	        Dispatch paragraph = null;  
	        Dispatch range = null;  
	        if(paragraphCount > paragraphsIndex && 0 < paragraphsIndex){    
	            paragraph = Dispatch.call(paragraphs, "Item", new Variant(paragraphsIndex)).toDispatch();  
	            range = Dispatch.get(paragraph, "Range").toDispatch();  
	            ret = Dispatch.get(range, "Text").toString();  
	        }     
	        return ret;  
	    }  
	  public static void setHeaderContent(Dispatch doc,String cont){  
	        Dispatch activeWindow = Dispatch.get(doc, "ActiveWindow").toDispatch();  
	        Dispatch view = Dispatch.get(activeWindow, "View").toDispatch();  
	        //Dispatch seekView = Dispatch.get(view, "SeekView").toDispatch();  
	        Dispatch.put(view, "SeekView", new Variant(9));         //wdSeekCurrentPageHeader-9  
	          
	        Dispatch headerFooter = Dispatch.get(doc, "HeaderFooter").toDispatch();  
	        Dispatch range = Dispatch.get(headerFooter, "Range").toDispatch();  
	        Dispatch.put(range, "Text", new Variant(cont));   
	        //String content = Dispatch.get(range, "Text").toString();  
	        Dispatch font = Dispatch.get(range, "Font").toDispatch();  
	          
	        Dispatch.put(font, "Name", new Variant("楷体_GB2312"));  
	        Dispatch.put(font, "Bold", new Variant(true));  
	        //Dispatch.put(font, "Italic", new Variant(true));  
	        //Dispatch.put(font, "Underline", new Variant(true));  
	        Dispatch.put(font, "Size", 9);  
	          
	        Dispatch.put(view, "SeekView", new Variant(0));         //wdSeekMainDocument-0恢复视图;  
	    }  
	  public static void main(String[] args) {  
		  WordToHtml.read("c://test//test.doc");
	  }
	    public static void main2(String[] args) {  
	        try {  
	           // String text = WordToHtml.readDoc("/Users/lijian/Documents/e.doc");  
	        	HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream("c://test//test.doc"));

	            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
	                    DocumentBuilderFactory.newInstance().newDocumentBuilder()
	                            .newDocument());
	            wordToHtmlConverter.processDocument(wordDocument);
	            org.w3c.dom.Document htmlDocument = wordToHtmlConverter.getDocument();
	            ByteArrayOutputStream out = new ByteArrayOutputStream();
	            DOMSource domSource = new DOMSource(htmlDocument);
	            StreamResult streamResult = new StreamResult(out);

	            TransformerFactory tf = TransformerFactory.newInstance();
	            Transformer serializer = tf.newTransformer();
	            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
	            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
	            serializer.setOutputProperty(OutputKeys.METHOD, "html");
	            serializer.transform(domSource, streamResult);
	            out.close();

	            String result = new String(out.toByteArray());
	            System.out.println(result);
//	            System.out.println(text);  
	        } catch (Exception e) {  
	            e.printStackTrace();  
	        }  
	          
	    }  	 
	    
}