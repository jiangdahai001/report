package com.example.report.util;

import org.dom4j.*;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.util.regex.Pattern.compile;

public class XMLConverter {
  public static final String TEMP_TAG = "geneplus_placeholder";
  public static void main(String[] args) throws Exception{
    System.out.println("XMLConverter");
    String sourcePath = "C:\\Users\\admin\\Desktop\\test\\ahello.xml";
    String targetPath = "C:\\Users\\admin\\Desktop\\test\\";
    convert(sourcePath, targetPath);
  }
  public static void convert(String sourcePath, String targetPath) throws Exception{
    // 读取xml
    SAXReader reader = new SAXReader();
    Document document = reader.read(new File(sourcePath));
    Element rootElement = document.getRootElement();

    // 以下处理都是开发过程测试用的，最后可以合并在一起
    handleWfldSimpleElement(rootElement);
    handleDividedDomain(rootElement);
    System.out.println("普通文本处理完成");
    handleTableForeach(rootElement);
    System.out.println("表格循环处理完成");
    handlePictureElement(rootElement, document);
    System.out.println("图片处理完成");
    document = removeTempTag(document);
    System.out.println("临时标签处理完成");
    writeXML(document, targetPath);
    System.out.println("保存xml完成");
  }
  // 将document写入到文件
  public static void writeXML(Document document,String targetPath) {
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
    String name = dateFormat.format(new Date());
    //将document对象内的内容 写入到文件里
    FileOutputStream file=null;
    OutputStreamWriter fos=null;
    XMLWriter writer=null;
    try {
      // 创建格式化类
      OutputFormat format = OutputFormat.createPrettyPrint();
      // 设置编码格式，默认UTF-8
      format.setEncoding("UTF-8");

      // 创建输出流文件不存在就创建
      file = new FileOutputStream(targetPath + name + ".xml");
      fos=new OutputStreamWriter(file, StandardCharsets.UTF_8);
      // 创建xml输出流
      writer = new XMLWriter(fos, format);
      // 生成xml文件
      writer.write(document);

    } catch (Exception e) {
      e.printStackTrace();
    }finally {
      try {
        writer.close();
        fos.close();
        file.close();
      } catch (IOException e) {
        e.printStackTrace();
      }
    }
  }

  /**
   * 合并域数据，无论是手动将docx另存为xml或者使用spiredoc的命令将docx另存为xml，都会出现域内容被分段的情况
   * 这里需要将一个域中完整的数据拼接到一起
   * 处理 w:fldChar 元素
   * <w:fldChar w:fldCharType="begin" />
   * <w:instrText xml:space="preserve" />
   * <w:fldChar w:fldCharType="separate" />
   * <w:t> 需要保留的内容 </w:t>
   * <w:fldChar w:fldCharType="end" />
   * @param element 目标元素
   */
  public static void handleDividedDomain(Element element) {
    // 遍历参数元素
    for (Iterator<Node> iterator = element.elementIterator(); iterator.hasNext(); ) {
      // 获取当前元素
      Element currentElement = (Element) iterator.next();
      if ("fldChar".equals(currentElement.getName()) && "begin".equals(currentElement.attributeValue("fldCharType"))) {
        Element parent = currentElement.getParent();
        Element grandpa = parent.getParent();
        Element domainElement = null;
        // 注意：这里的Element.indexOf获取的索引和Element.elements()获取的list的索引不能一一对应，目前看是x=(n-1)/2
        int indexStart = (grandpa.indexOf(parent) - 1)/2;
        int indexEnd = 0;
        StringBuffer domainValue = new StringBuffer();
        // 找到分割域的起始和结束的索引，并将所有索引内容拼接到一起
        for (int i = indexStart; i < grandpa.elements().size(); i++) {
          Element el = (Element) grandpa.elements().get(i);
          if (el.element("t") != null) {
            domainValue.append(el.element("t").getTextTrim());
            if (domainElement == null) {
              domainElement = (Element) el.clone();
            }
          }
          if (el.element("fldChar") != null) {
            if ("end".equals(el.element("fldChar").attributeValue("fldCharType"))) {
              indexEnd = i;
              break;
            }
          }
        }
        // 删除分割的域内容
        for (int i = indexEnd; i >= indexStart; i--) {
          ((Element) grandpa.elements().get(i)).detach();
        }
        // 在之前分割域开始的位置加上新的拼接好的域元素
        domainElement.element("t").setText(domainValue.toString());
        grandpa.elements().add(indexStart, domainElement);
      }
      handleDividedDomain(currentElement);
    }
  }
  /**
   * 处理 w:fldSimple 元素
   * <w:fldSimple>需要保留的内容</w:fldSimple>
   * @param element 节点元素
   */
  public static void handleWfldSimpleElement(Element element) {
    for (Iterator<Element> iterator = element.elementIterator(); iterator.hasNext(); ) {
      Element el = iterator.next();
      if("fldSimple".equals(el.getName())) {
        Element wr = (Element) el.element("r").clone();
//        int index = (el.getParent().indexOf(el) - 1) / 2;
        int index = el.getParent().indexOf(el);
        el.getParent().elements().add(index, wr);
        el.getParent().remove(el);
      }
      handleWfldSimpleElement(el);
    }
  }
  /**
   * 处理table中的foreach循环
   * <w:tbl>
   *   <w:tr>
   *     <w:tc>
   *       <w:p>
   *         <w:r>
   *           <w:t>#foreach / #end</w:t>
   *         </w:r>
   *       </w:p>
   *     </w:tc>
   *   </w:tr>
   * </w:tbl>
   * @param element
   */
  public static void handleTableForeach(Element element) {
    for (Iterator<Element> iterator = element.elementIterator(); iterator.hasNext(); ) {
      Element el = iterator.next();
      if("t".equals(el.getName()) && (el.getText().contains("#foreach") || el.getText().contains("#end"))) {
        Element tbl = el.getParent().getParent().getParent().getParent().getParent();
        Element tr = el.getParent().getParent().getParent().getParent();
//        int index = (tbl.indexOf(tr) - 1) / 2;
        int index = tbl.indexOf(tr);
        Element foreach = DocumentHelper.createElement(TEMP_TAG);
        foreach.setText(el.getText());
        tbl.elements().add(index, foreach);
        tbl.remove(tr);
      }
      handleTableForeach(el);
    }
  }

  /**
   * 将临时标签删除，只保留标签中的内容
   * @param document 文档对象
   * @return 返回修改后的文档对象
   * @throws Exception 异常
   */
  public static Document removeTempTag(Document document) throws Exception {
    String s = document.asXML();
    Pattern compile = compile("<" + TEMP_TAG + ".*?>");
    Matcher matcher = compile.matcher(s);
    while (matcher.find()) {
      s = s.replace(matcher.group(), "");
    }
    s = s.replaceAll("</"+ TEMP_TAG +">", "");
    return DocumentHelper.parseText(s);
  }

  /**
   * 处理占位图片元素，
   * 0，在word中用”替换文字“用域来标记图片，
   * 1，在xml中通过 w:drawing 标签下 wp:docPr 的 descr 属性的值来找到对应的域
   * 2，找到对应的a:blip 标签的 r:embed 属性值，就是word给图片分配的id
   * 3，通过id属性找到对应的Relationship，获取Target属性的值，就是对应图片在word中的实际名称
   * 4，通过名称（前面要拼上/word/）找到对应图片的base64编码，将编码内容替换为域名
   * 5，将前面 wp:docPr 的descr属性清空，将pic:cNvPr 的descr属性清空
   * @param element 目标元素
   * @param document 文档对象
   */
  public static void handlePictureElement(Element element, Document document) {
    for (Iterator<Element> iterator = element.elementIterator(); iterator.hasNext(); ) {
      Element el = iterator.next();
      if("docPr".equals(el.getName())) {
        String domainName = el.attributeValue("descr");
        if(domainName!=null && domainName.contains("{") && domainName.contains("}")) {
          Element inline = el.getParent();
          String rId = inline.element("graphic").element("graphicData").element("pic").element("blipFill").element("blip").attributeValue("embed");
          System.out.println("rId: " + rId);
//          List<String> targetList = new ArrayList<>();
//          getPictureTargetById(rId, document.getRootElement(), targetList);
//          String target = targetList.get(0);
//          System.out.println("target: " + target);
//          List<Element> partList = new ArrayList<>();
//          getPkgPartByName(target, document.getRootElement(), partList);
//          Element part = partList.get(0);
//          part.element("binaryData").setText(domainName);
          Element targetElement = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.openxmlformats.org/package/2006/relationships' and local-name()='Relationship' and @Id='"+rId+"']");
          String target = targetElement.attributeValue("Target");
          Element partElement = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.microsoft.com/office/2006/xmlPackage' and local-name()='part' and @pkg:name='/word/"+target+"']");
          partElement.element("binaryData").setText(domainName);

          // 将 wp:docPr 的 descr 属性删除
          Attribute docPrDescr = el.attribute("descr");
          docPrDescr.detach();
          // 将 pic:cNvPr 的 descr 属性删除
          Attribute piccNvPrDescr =inline.element("graphic").element("graphicData").element("pic").element("nvPicPr").element("cNvPr").attribute("descr");
          piccNvPrDescr.detach();
        }
      }
      handlePictureElement(el, document);
    }
  }
  private static void getPkgPartByName(String name, Element element, List<Element> partList) {
    for(Iterator<Element> iterator = element.elementIterator();iterator.hasNext();) {
      Element el = iterator.next();
      if("part".equals(el.getName()) && el.attributeValue("name").equals("/word/" + name)) {
        partList.add(el);
      }
      getPkgPartByName(name, el, partList);
    }
  }
  private static void getPictureTargetById(String rId, Element element, List<String> targetList) {
    for (Iterator<Element> iterator = element.elementIterator(); iterator.hasNext(); ) {
      Element el = iterator.next();
      if("Relationship".equals(el.getName()) && rId.equals(el.attributeValue("Id"))) {
        targetList.add(el.attributeValue("Target"));
      }
      getPictureTargetById(rId, el, targetList);
    }
  }
}