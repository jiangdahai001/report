package com.example.report.util;

import org.dom4j.*;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class Dom4jTest {
  public final static String PATH = "C:\\Users\\admin\\Desktop\\test\\";
  public static void main(String[] args) throws Exception {
    System.out.println("hello, world");
    String path = "C:\\Users\\admin\\Desktop\\test\\2023-06-15-14-57-52.xml";
    testXPath(path);
  }

  // 创建xml
  public static void createXML() {
    // 创建文档对象。
    Document document = DocumentHelper.createDocument();
    // 文档增加第一个节点，即根节点，一个文档只能有一个根节点，多加出错  然后创建的其他节点都是在根节点内
    Element root = document.addElement("xml_gen");
    // 添加注释 (可有可无)   注释是在节点上方 的      <!--?? -->
    root.addComment("第一个技能");
    // 根节点下添加节点
    Element first = root.addElement("day_1");
    // 节点添加属性
    first.addAttribute("name", "独孤九剑");
    // 节点下添加节点
    Element info = first.addElement("info");
    // 节点设置内容数据
    info.setText("为独孤求败所创，变化万千，凌厉无比。其传人主要有风清扬、令狐冲。");
    // 同上  增加其他节点，内容，属性等
    Element second = root.addElement("day_2");
    second.addAttribute("name", "葵花宝典");
    Element info2 = second.addElement("info");
    info2.setText("宦官所创，博大精深，而且凶险至极。练宝典功夫时，首先要自宫净身。");
    // 创建节点
    Element third = DocumentHelper.createElement("skill");
    // 将节点加入到根节点中
    root.add(third);
    // 创建属性
    Attribute name = DocumentHelper.createAttribute(third, "name", "北冥神功");
    // 将属性加入到节点上
    third.add(name);
    // 创建子节点并加入到节点中
    Element info3 = DocumentHelper.createElement("day_1");
    info3.setText("逍遥派的顶级内功之一，能吸人内力转化为自己所有，威力无穷。");
    third.add(info3);
    writeXML(document);
  }
  // 读取xml
  public static Document readXML(String path) throws Exception{
    SAXReader reader = new SAXReader();
    Document document = reader.read(new File(path));
    Element rootElement = document.getRootElement();
    traversalElement(rootElement);
    return document;
  }
  // 遍历element下的所有元素
  public static void traversalElement(Element element) {
    for(Iterator<Element> iterator=element.elementIterator();iterator.hasNext();) {
      Element el = iterator.next();
      System.out.println("elementName: " + el.getName() + "--elementValue: " + el.getTextTrim());
      List<Attribute> attributeList = el.attributes();
      for(Attribute attribute:attributeList) {
        System.out.println("elementAttributeName: " + attribute.getName() + "--elementAttributeValue: " + attribute.getValue());
      }
      traversalElement(el);
    }
  }
  // 解析字符串
  public static void parseString() throws DocumentException {
    String xml="<messge>"+
      "<name>hu</name>"+
      "<age>22</age>" +
      "</messge>";
    Document document= DocumentHelper.parseText(xml);
    List list = document.selectNodes("/messge");
    Element message = (Element) list.get(0);
    System.out.println(message.getStringValue().trim());
    // 查询 document的内容
    System.out.println("原xml");
    String s = document.asXML();
    System.out.println(s);
  }
  // 将document写入到文件
  public static void writeXML(Document document) {
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
      file = new FileOutputStream(PATH + name + ".xml");
      fos=new OutputStreamWriter(file,"utf-8");
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
  // 修改xml
  public static void modifyXML(String path) throws Exception {
    Document document = readXML(path);
    List list = document.selectNodes("/xml_gen/day_1/info");
    for(Iterator iterator = list.iterator(); iterator.hasNext();) {
      Element info = (Element) iterator.next();
      info.setText("new info");
      info.setName("info_new");
      info.addAttribute("num", "101shengong");
      if(info.getParent().remove(info)) {
        System.out.println("删除成功！");
      }
    }
    List list1 = document.selectNodes("/xml_gen/day_2");
    for(Iterator iterator = list1.iterator();iterator.hasNext();) {
      Element day2 = (Element) iterator.next();
      System.out.println(day2.getName());
      Element newElement = day2.addElement("newElement");
      newElement.setText("this is new add element");
      newElement.addAttribute("language", "cn");
      newElement.add(DocumentHelper.createText("xxxx"));
    }
    writeXML(document);
  }
  // 为xml的element添加兄弟元素
  public static void addSibling(String path) throws Exception {
    Document document = readXML(path);
    List list = document.selectNodes("/xml_gen/day_2");
    int index = 0;
    Element day2=null;
    for(Iterator iterator = list.iterator(); iterator.hasNext();) {
      day2 = (Element) iterator.next();
      index = day2.getParent().elements().indexOf(day2);
    }
    Element sibling = DocumentHelper.createElement("sibling");
    sibling.setText("abc");
    day2.getParent().elements().add(index,sibling);
    writeXML(document);
  }
  // 去掉外层元素，主要处理 w:fldSimple 标签
  public static void getRidOfOuterLayer(Document document) throws Exception {
    List list = document.selectNodes("/xml_gen/day_2");
    int index = 0;
    Element day2 = null;
    Element parent = null;
    for(Iterator iterator = list.iterator(); iterator.hasNext();) {
      day2 = (Element) iterator.next();
      parent = day2.getParent();
      index = parent.elements().indexOf(day2);
    }
    Element sibling = (Element) day2.element("info").clone();
    parent.elements().add(index, sibling);
    parent.remove(day2);
    writeXML(document);
  }

  public static void testXPath(String path) throws Exception {
    SAXReader reader = new SAXReader();
    Document document = reader.read(new File(path));
    String name = "独孤九剑";
    name = "葵花宝典";
    Element Relationship = (Element) document.selectSingleNode("//day[@name='"+name+"']");
    String target = Relationship.attributeValue("att1");
    System.out.println("target: " + target);
  }
}
