package com.example.report.util;

import com.spire.doc.FileFormat;
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
import java.util.*;
import java.util.function.ToDoubleBiFunction;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static java.util.regex.Pattern.compile;

/**
 * 从编写好的模板docx，另存为或者spiredoc转换成为xml文件，然后转换为可以给velocity填值的xml文件
 */
public class XMLConverter {
  public static final String TEMP_TAG = "geneplus_placeholder";
  // 注意：这里的Element.indexOf获取的索引和Element.elements()获取的list的索引不能一一对应，目前看是x=(n-1)/2
  public static final boolean indexFitFlag = false;
  public static void main(String[] args) throws Exception{
    System.out.println("XMLConverter");
    String docxPath = "C:\\Users\\admin\\Desktop\\test\\ahello.docx";
    String sourcePath = "C:\\Users\\admin\\Desktop\\test\\ahello.xml";
    String targetPath = "C:\\Users\\admin\\Desktop\\test\\";

    //加载模板
    com.spire.doc.Document document = new com.spire.doc.Document();
    document.loadFromFile(docxPath, FileFormat.Docx);
    //将模板另存为xml
    document.saveToFile(sourcePath, FileFormat.Word_Xml);

    convert(sourcePath, targetPath);
  }
  public static void convert(String sourcePath, String targetPath) throws Exception{
    // 读取xml
    SAXReader reader = new SAXReader();
    Document document = reader.read(new File(sourcePath));
    // 处理简单域元素
    handleWfldSimpleElement(document);
    // 合并被拆分的域
    handleDividedDomain(document);
    System.out.println("简单域、拆分域合并处理完成");
    // 处理独立段落中velocity标签语法
    handleVelocityParagraphTag(document);
    // 处理段落中行内velocity标签语法
    handleVelocityInlineTag(document);
    // 处理行内颜色参数
    handleInlineColor(document);
    // 处理段落的行内循环
    handleInlineForeach(document);
    // 处理table中行合并
    handleTableVmerge(document);
    // 处理table中列合并
    handleTableGridSpan(document);
    // 处理table中foreach循环
    handleTableForeach(document);
    System.out.println("表格循环处理完成");
    // 处理table中cell的背景色
    handleTableTcShading(document);
    // 处理占位图片
    handlePictureElement(document);
    // 处理图片循环
    handlePictureForeach(document);
    handlePictureInlineForeach(document);
    removeRedundantPictureAttribute(document);
    System.out.println("图片处理完成");
    // 移除临时标签，保留标签的text
    document = removeTempTag(document);
    System.out.println("临时标签处理完成");
    // 保存新的xml文件
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
   * @param document 文档元素
   */
  public static void handleDividedDomain(Document document) {
//    List<Element> beginList = document.selectNodes("//*[local-name()='fldChar' and @fldCharType='begin']");
    List<Element> beginList = document.selectNodes("//*[local-name()='fldChar']");
    for(Element wt:beginList) {
      if("begin".equals(wt.attributeValue("fldCharType"))) {
        Element wr = (Element) wt.selectSingleNode("ancestor::w:r");
        Element wrParent = wr.getParent();
        Element domainElement = null;
        // 注意：这里的Element.indexOf获取的索引和Element.elements()获取的list的索引不能一一对应，目前看是x=(n-1)/2
        int index = wrParent.indexOf(wr);
        if (indexFitFlag) index = (index - 1) / 2;
        StringBuffer domainValue = new StringBuffer();
        List<Element> domainList = wr.selectNodes("following-sibling::w:r");
        for (Element domain : domainList) {
          domain.detach();
          if (domain.element("t") != null) {
            domainValue.append(domain.element("t").getTextTrim());
            if (domainElement == null) {
              domainElement = (Element) domain.clone();
            }
          }
          if (domain.element("fldChar") != null) {
            if ("end".equals(domain.element("fldChar").attributeValue("fldCharType"))) {
              break;
            }
          }
        }
        // 在之前分割域开始的位置加上新的拼接好的域元素
        domainElement.element("t").setText(domainValue.toString());
        wrParent.elements().add(index, domainElement);
        wr.detach();
      }
    }
  }
  /**
   * 处理 w:fldSimple 元素
   * <w:fldSimple>需要保留的内容</w:fldSimple>
   * @param document 文档元素
   */
  public static void handleWfldSimpleElement(Document document) {
    List<Element> wfldSimpleList = document.selectNodes("//*[local-name()='fldSimple']");
    for(Element element:wfldSimpleList) {
      Element wr = (Element) element.selectSingleNode("descendant::w:r").clone();
      int index = element.getParent().indexOf(element);
      if(indexFitFlag) index = (index - 1) / 2;
      element.getParent().elements().add(index, wr);
      element.getParent().remove(element);
    }
  }
  /**
   * 处理在独立段落中，使用velocity的set、if、foreach、macro等语法
   * 前提是都是独立的段落中，即语法所在的父元素wp是后面要被移除的
   * @param document 文档对象
   */
  public static void handleVelocityParagraphTag(Document document) {
    StringBuffer sb = new StringBuffer();
    // if 标签相关
    sb.append("contains(text(),'#p_if')");
    sb.append(" or contains(text(), '#p_elseif')");
    sb.append(" or contains(text(), '#p_else')");
    sb.append(" or contains(text(), '#p_end')");
    // set 标签相关
    sb.append(" or contains(text(), '#p_set')");
    // macro 标签相关
    sb.append(" or contains(text(), '#p_macro')");
    // foreach 标签相关，p_end在上面if标签有了
    sb.append(" or contains(text(), '#p_foreach')");
    List<Element> pifList = document.selectNodes("//*[local-name()='t' and ("+sb.toString()+")]");
    for(Element wt:pifList) {
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      Element wpParent = wp.getParent();
      int index = wpParent.indexOf(wp);
      if(indexFitFlag) index = (index - 1) / 2;
      Element tmp = DocumentHelper.createElement(TEMP_TAG);
      String text = wt.getText().replaceAll("^#p_", "#");
      tmp.setText(text);
      wpParent.elements().add(index, tmp);
      wpParent.remove(wp);
    }
  }

  /**
   * 处理在段落行内中，使用velocity的set、if、foreach、macro等语法
   * 前提是都是段落行内中，即语法所在的父元素wr是后面要被移除的
   * @param document 文档对象
   */
  public static void handleVelocityInlineTag(Document document) {
    StringBuffer sb = new StringBuffer();
    // if 标签相关
    sb.append("contains(text(),'#inline_if')");
    sb.append(" or contains(text(), '#inline_elseif')");
    sb.append(" or contains(text(), '#inline_else')");
    sb.append(" or contains(text(), '#inline_end')");
    // set 标签相关
    sb.append(" or contains(text(), '#inline_set')");
    // macro 标签相关
    sb.append(" or contains(text(), '#inline_macro')");
    // foreach 标签相关，inline_end在上面if标签有了
    sb.append(" or contains(text(), '#inline_foreach')");
    List<Element> elementList = document.selectNodes("//*[local-name()='t' and ("+sb.toString()+")]");
    for(Element wt:elementList) {
      Element wr = (Element) wt.selectSingleNode("ancestor::w:r");
      Element wrParent = wr.getParent();
      int index = wrParent.indexOf(wr);
      if(indexFitFlag) index = (index - 1) / 2;
      Element tmp = DocumentHelper.createElement(TEMP_TAG);
      String text = wt.getText().replaceAll("^#inline_", "#");
      tmp.setText(text);
      wrParent.elements().add(index, tmp);
      wrParent.remove(wr);
    }
  }

  /**
   * 处理行内字体颜色，包括字体的颜色，高亮，底纹（color，highlight，shading）
   * #inline_color_begin(color="FF0000",highlight="yellow",shading="0000FF")实际内容#inline_color_end
   * @param document 文档元素
   */
  public static void handleInlineColor(Document document) {
    StringBuffer sb = new StringBuffer();
    sb.append("contains(text(),'#inline_color_begin')");
    List<Element> elementList = document.selectNodes("//*[local-name()='t' and ("+sb.toString()+")]");
    for(Element wt:elementList) {
      Element wr = (Element) wt.selectSingleNode("ancestor::w:r");
      String content = wt.getText().replaceAll("^#inline_color_begin\\(|\\)$", "");
      Map<String, String> paramsMap = getParamsMap(content);
      List<Element> siblingList = wr.selectNodes("following-sibling::*");
      for(Element sibling:siblingList) {
        if(sibling.getStringValue().contains("#inline_color_end")) {
          sibling.detach();
          break;
        }
        // 设置字体颜色
        if(paramsMap.containsKey("color")) {
          Element wcolor = (Element) sibling.selectSingleNode("descendant::*/w:rPr/w:color");
          if (wcolor == null) {
            Element color = DocumentHelper.createElement("w:color");
            color.addAttribute("w:val", paramsMap.get("color"));
            sibling.element("rPr").add(color);
          } else {
            wcolor.addAttribute("w:val", paramsMap.get("color"));
          }
        }
        // 设置字体高亮色
        if(paramsMap.containsKey("highlight")) {
          Element whighlight = (Element) sibling.selectSingleNode("descendant::*/w:rPr/w:highlight");
          if (whighlight == null) {
            Element highlight = DocumentHelper.createElement("w:highlight");
            highlight.addAttribute("w:val", paramsMap.get("highlight"));
            sibling.element("rPr").add(highlight);
          } else {
            whighlight.addAttribute("w:val", paramsMap.get("highlight"));
          }
        }
        // 设置字体底纹
        if(paramsMap.containsKey("shading")) {
          Element wshd = (Element) sibling.selectSingleNode("descendant::*/w:rPr/w:shd");
          if (wshd == null) {
            Element shd = DocumentHelper.createElement("w:shd");
            shd.addAttribute("w:fill", paramsMap.get("shading"));
            shd.addAttribute("w:val", "clear");
            shd.addAttribute("w:color", "auto");
            sibling.element("rPr").add(shd);
          } else {
            wshd.addAttribute("w:fill", paramsMap.get("shading"));
            wshd.addAttribute("w:val", "clear");
            wshd.addAttribute("w:color", "auto");
          }
        }
      }
      wr.detach();
    }
  }

  /**
   * 将参数内容转换为参数map，方便参数使用
   * content: param1="asdf",param2="dsdfds"
   * @param content 参数文本
   * @return 参数map
   */
  private static Map<String, String> getParamsMap(String content) {
    String[] paramsArray = content.split(",");
    Map<String, String> paramsMap = new HashMap<>();
    Arrays.stream(paramsArray).forEach(param -> {
      String[] tmp = param.split("=");
      String key = tmp[0];
      String value = tmp[1].replaceAll("\"", "");
      paramsMap.put(key, value);
    });
    return paramsMap;
  }

  /**
   * 处理table的行合并
   * 模板中使用#tbl_vmerge(开始合并的条件)
   * 下面代码自动将开始合并的条件放到#if中，如果满足则添加<w:vMerge w:val="restart"/>
   * 如果不满足则添加<w:vMerge />
   * @param document 文档元素
   */
  public static void handleTableVmerge(Document document) {
    List<Element> vmergeList = document.selectNodes("//*[local-name()='t' and contains(text(), '#tbl_vmerge')]");
    for(Element wt:vmergeList) {
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      Element tc = (Element) wp.selectSingleNode("ancestor::w:tc");
      String domainValue = wt.getText().replaceAll("^#tbl_vmerge\\(|\\)$", "");
      Element vmerge = DocumentHelper.createElement(TEMP_TAG);
      vmerge.setText("#if(" + domainValue + ")");
      Element start = DocumentHelper.createElement("w:vMerge");
      start.addAttribute("w:val", "restart");
      Element velse = DocumentHelper.createElement(TEMP_TAG);
      velse.setText("#else");
      Element end = DocumentHelper.createElement("w:vMerge");
      Element vend = DocumentHelper.createElement(TEMP_TAG);
      vend.setText("#end");

      vmerge.add(start);
      vmerge.add(velse);
      vmerge.add(end);
      vmerge.add(vend);

      tc.element("tcPr").elements().add(vmerge);
      wp.detach();
    }
  }

  /**
   * 处理table的列合并
   * 模板中使用#tbl_grid_span(合并的列数)
   * 如果列数大于1，则添加w:gridSpan元素，并配置w:val的值
   * 如果列数为1，则不添加w:gridSpan
   * 如果列数为0，则删除当前的tc
   * @param document 文档元素
   */
  public static void handleTableGridSpan(Document document) {
    List<Element> gridSpanList = document.selectNodes("//*[local-name()='t' and contains(text(), '#tbl_grid_span')]");
    for(Element wt:gridSpanList) {
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      Element tc = (Element) wp.selectSingleNode("ancestor::w:tc");
      Element tr = (Element) tc.selectSingleNode("ancestor::w:tr");
      String domainValue = wt.getText().replaceAll("^#tbl_grid_span\\(|\\)$", "");
      Element tcPrefix = DocumentHelper.createElement(TEMP_TAG);
      tcPrefix.setText("#set($tmp="+domainValue+") #set($val=$tmp.trim()) #if($number.toNumber($val) gt 0)");
      wp.detach();
      Element newTc = (Element) tc.clone();
      Element tcSuffix = DocumentHelper.createElement(TEMP_TAG);
      tcSuffix.setText("#end");

      Element gridSpan = DocumentHelper.createElement("w:gridSpan");
      gridSpan.addAttribute("w:val", "$!{val}");
      newTc.element("tcPr").elements().add(gridSpan);

      int index = tr.indexOf(tc);
      if(indexFitFlag) index = (index - 1) / 2;
      // 注意这里要先加后面的元素，因为index不变，是往前插入的
      tr.elements().add(index, tcSuffix);
      tr.elements().add(index, newTc);
      tr.elements().add(index, tcPrefix);
      tc.detach();
    }
  }

  /**
   * 处理table中tr级别的foreach循环
   * @param document 文档元素
   */
  public static void handleTableForeach(Document document) {
    // 处理tr级别的foreach循环
    // tbl_tr_foreach, tbl_tr_end
    // tbl_tr_if, tbl_tr_else, tbl_tr_elseif, tbl_tr_end
    StringBuffer sb = new StringBuffer();
    sb.append("contains(text(),'#tbl_tr_foreach')");
    sb.append(" or contains(text(), '#tbl_tr_end_foreach')");
    sb.append(" or contains(text(), '#tbl_tr_if')");
    sb.append(" or contains(text(), '#tbl_tr_else')");
    sb.append(" or contains(text(), '#tbl_tr_elseif')");
    sb.append(" or contains(text(), '#tbl_tr_end')");
    sb.append(" or contains(text(), '#tbl_tr_set')");
    // 处理循环图片相关的内容，这里只处理图片，标签相关内容在下面处理
    List<Element> trForeachList = document.selectNodes("//*[local-name()='t' and (contains(text(), '#tbl_tr_foreach'))]");
    for(Element wt:trForeachList) {
      String foreachContent = wt.getText().replaceAll("^#tbl_tr_", "#").replaceAll("_foreach$", "");
      Element tr = (Element) wt.selectSingleNode("ancestor::w:tr");
      List<Element> siblingList = tr.selectNodes("following-sibling::*");
      for(Element sibling:siblingList) {
        if(sibling.getStringValue().contains("#tbl_tr_end_foreach")) break;
        List<Element> drawingList = sibling.selectNodes("descendant::*//w:drawing");
        for (Element drawing : drawingList) {
          Element docPr = (Element) drawing.selectSingleNode("descendant::wp:docPr");
          String domainName = docPr.attributeValue("descr");
          if(domainName == null) throw new RuntimeException("图片没有填写”替换文字“，请检查");
          // 如果时占位图片或者固定图片，不在这里处理
          if (domainName.matches("^placeholder.*?|^static.*?")) continue;
          addPictureElement(document, drawing, foreachContent, domainName);
        }
      }
    }
    // 处理标签相关内容
    List<Element> trList = document.selectNodes("//*[local-name()='t' and ("+sb.toString()+")]");
    for(Element wt:trList) {
      Element tr = (Element) wt.selectSingleNode("ancestor::w:tr");
      Element tbl = (Element) wt.selectSingleNode("ancestor::w:tbl");
      int index = tbl.indexOf(tr);
      if(indexFitFlag) index = (index - 1) / 2;
      Element foreach = DocumentHelper.createElement(TEMP_TAG);
      String text = wt.getText().replaceAll("^#tbl_tr_", "#").replaceAll("_foreach$", "");
      foreach.setText(text);
      tbl.elements().add(index, foreach);
      tbl.remove(tr);
    }
  }

  /**
   * 处理table中tc的shading设置
   * 模板中语法:#tbl_tc_shading(val), val 为颜色值，例如C9C9C9
   * 注意如果使用$render.eval，则返回值不能有引号
   * 如果val直接为颜色值，则需要用引号包裹
   * @param document 文档元素
   */
  public static void handleTableTcShading(Document document) {
    List<Element> tcShadingList = document.selectNodes("//*[local-name()='t' and contains(text(), '#tbl_tc_shading')]");
    for(Element wt:tcShadingList) {
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      Element tc = (Element) wp.selectSingleNode("ancestor::w:tc");
      String wtFill = wt.getText().replaceAll("^#tbl_tc_shading\\(|\\)$", "");
      Element shdPrefix = DocumentHelper.createElement(TEMP_TAG);
      shdPrefix.setText("#set($fill="+wtFill+")");
      Element shd = DocumentHelper.createElement("w:shd");
      shd.addAttribute("w:color", "auto");
      shd.addAttribute("w:fill","$!{fill}");
      shd.addAttribute("w:val", "clear");
      tc.element("tcPr").elements().add(shdPrefix);
      tc.element("tcPr").elements().add(shd);
      wp.detach();
    }
  }
  /**
   * 处理段落中的行内循环
   * #p开头的标签是要单独写一行的
   * #inline开头的标签是同其他内容写在同一行内的
   * @param document 文档元素
   */
  public static void handleInlineForeach(Document document) {
    // 处理行内foreach循环
    List<Element> inlineForeachList = document.selectNodes("//*[local-name()='t' and contains(text(), '#p_inline_foreach')]");
    for(Element wt:inlineForeachList) {
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      List<Element> wpSiblingList = wp.selectNodes("following-sibling::w:p");
      // 将所有wp后续的兄弟元素中的wr子元素都放到第一个wp中，其余的wp兄弟元素就可以detach了，直到遇到#p_inline_end_foreach
      // 如果有临时元素也要放到第一个wp中，
      boolean endForeach = false;
      for(Element sibling:wpSiblingList) {
        if("#p_inline_end_foreach".equals(sibling.getStringValue())) {
          endForeach = true;
        }
        List<Element> wrList = sibling.selectNodes("descendant::w:r|descendant::"+TEMP_TAG+"");
        wrList.forEach(wr -> {
          wr.detach();
        });
        wp.elements().addAll(wrList);
        sibling.detach();
        if(endForeach) break;
      }
      // 将wr元素中，自己写的velocity语法使用临时标签包裹，后面再去掉临时标签，只保留text
      StringBuffer tcInlineBuffer = new StringBuffer();
      tcInlineBuffer.append("contains(text(),'#p_inline_foreach')");
      tcInlineBuffer.append(" or contains(text(), '#p_inline_end_foreach')");
      tcInlineBuffer.append(" or contains(text(), '#p_inline_if')");
      tcInlineBuffer.append(" or contains(text(), '#p_inline_else')");
      tcInlineBuffer.append(" or contains(text(), '#p_inline_elseif')");
      tcInlineBuffer.append(" or contains(text(), '#p_inline_end')");
      List<Element> wtList = wp.selectNodes("descendant::*[local-name()='t' and ("+tcInlineBuffer.toString()+")]");
      for(Element wt2:wtList) {
        Element wr = (Element) wt2.selectSingleNode("ancestor::w:r");
        int index = wp.indexOf(wr);
        if(indexFitFlag) index = (index - 1) / 2;
        Element foreach = DocumentHelper.createElement(TEMP_TAG);
        String text = wt2.getText().replaceAll("^#p_inline_", "#").replaceAll("_foreach$", "");
        foreach.setText(text);
        wp.elements().add(index, foreach);
        wp.remove(wr);
      }
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
   * 5，将前面 wp:docPr 的descr属性删除，将pic:cNvPr 的descr属性删除
   * @param document 文档对象
   */
  public static void handlePictureElement(Document document) {
    List<Element> drawingList = document.selectNodes("//*[local-name()='drawing']");
    for(Element drawing:drawingList) {
      Element docPr = (Element) drawing.selectSingleNode("descendant::wp:docPr");
      String domainName = docPr.attributeValue("descr");
      // 这里只处理占位图片，其他的图片不处理
      if(!domainName.matches("^placeholder.*?")) continue;
      domainName = docPr.attributeValue("descr").replaceAll("^placeholder", "");
      String rId = ((Element) drawing.selectSingleNode("descendant::*[namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main' and local-name()='blip']")).attributeValue("embed");
      Element targetElement = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.openxmlformats.org/package/2006/relationships' and local-name()='Relationship' and @Id='"+rId+"']");
      String target = targetElement.attributeValue("Target");
      Element partElement = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.microsoft.com/office/2006/xmlPackage' and local-name()='part' and @pkg:name='/word/"+target+"']");
      partElement.element("binaryData").setText(domainName);
    }
  }

  /**
   * 处理换行图片foreach循环，单行多个或者多个都可以，需要在图片的“替换文字”处提供对应的域明，用于标识不同图片
   * 找到#pic_foreach, #pic_end标签，在#foreach和#end中间，所有w:drawing元素就是图片元素
   * 替换图片元素中的rId内容，引入$!{foreach.index}实现图片数量的动态变化
   * 新增Relationship及pkg:part，同样引入$!{foreach.index}实现图片数量的动态变化
   * 换行循环在图片的wp元素前后加foreach, end语法
   * @param document 目标文档
   */
  public static void handlePictureForeach(Document document) {
    StringBuffer sb = new StringBuffer();
    sb.append("contains(text(),'#pic_foreach')");
    sb.append(" or contains(text(), '#pic_end')");
    List<Element> picForeachList = document.selectNodes("//*[namespace-uri()='http://schemas.openxmlformats.org/wordprocessingml/2006/main' and local-name()='t' and ("+sb.toString()+")]");
    for (Element wt:picForeachList) {
      String foreachContent = wt.getTextTrim().replaceAll("^#pic_", "#");
      // 找到w:p祖先元素
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      if(wt.getTextTrim().contains("foreach")) {
        List<Element> wDrawingList = new ArrayList<>();
        List<Element> wpSiblingList = wp.selectNodes("following-sibling::*");
        for(Element sibling:wpSiblingList) {
          String content = sibling.getStringValue();
          if(content.contains("#pic_end")) break;
          wDrawingList.addAll(sibling.selectNodes("descendant::w:drawing"));
        }
        // 放占位图片的w:p中可能不止一个占位图片，需要使用他们的“替代文字”来标识不同的图片
        handleDrawingList(wDrawingList, document, foreachContent);
      }
      // 将w:p祖先元素替换成临时标签元素
      int index = wp.getParent().indexOf(wp);
      if(indexFitFlag) index = (index - 1) / 2;
      Element tmp = DocumentHelper.createElement(TEMP_TAG);
      String text = wt.getText().replaceAll("^#pic_", "#");
      tmp.setText(text);
      wp.getParent().elements().add(index, tmp);
      wp.detach();
    }
  }
  /**
   * 处理行内图片foreach循环，单行多个或者多个都可以，需要在图片的“替换文字”处提供对应的域明，用于标识不同图片
   * 找到#pic_inline_foreach, #pic_inline_end标签，在#foreach和#end中间，所有w:drawing元素就是图片元素
   * 替换图片元素中的rId内容，引入$!{foreach.index}实现图片数量的动态变化
   * 新增Relationship及pkg:part，同样引入$!{foreach.index}实现图片数量的动态变化
   * 行内循环在图片的wp元素内部加foreach, end语法
   * @param document 目标文档
   */
  public static void handlePictureInlineForeach(Document document) {
    StringBuffer sb = new StringBuffer();
    sb.append("contains(text(),'#pic_inline_foreach')");
    sb.append(" or contains(text(), '#pic_inline_end')");
    List<Element> picForeachList = document.selectNodes("//*[namespace-uri()='http://schemas.openxmlformats.org/wordprocessingml/2006/main' and local-name()='t' and ("+sb.toString()+")]");
    for (Element wt:picForeachList) {
      // 获取foreach标签内容
      String foreachContent = wt.getTextTrim().replaceAll("^#pic_inline_", "#");
      // 找到w:p祖先元素，需要把foreach放到接下来目标w:p第一个wt位置中，把end放到目标w:p最后一个wt位置
      Element wp = (Element) wt.selectSingleNode("ancestor::w:p");
      if(foreachContent.contains("foreach")) {
        // 放占位图片的w:p中可能不止一个占位图片，需要使用他们的“替代文字”来标识不同的图片
        List<Element> wDrawingList = wp.selectNodes("descendant::w:drawing");
        handleDrawingList(wDrawingList, document, foreachContent);
      }
      // 行内循环，将临时标签放到pic所在的w:p元素中，foreach放在w:pPr后面，end放在最后
      Element tmp = DocumentHelper.createElement(TEMP_TAG);
      String text = wt.getText().replaceAll("^#pic_inline_", "#");
      tmp.setText(text);
      if(text.contains("foreach")) {
        // 如果是pic_inline_multi_foreach，则将foreach语句放到图片所在w:p的w:pPr后面
        int index = wp.indexOf(wp.element("pPr"));
        if(indexFitFlag) index = (index - 1) / 2;
        wp.elements().add(index + 1, tmp);
      } else {
        // 如果是pic_inline_end，则将end语句放到图片所在w:p的最后即可
        wp.elements().add(tmp);
      }
      wt.detach();
    }
  }
  private static void handleDrawingList(List<Element> wDrawingList, Document document, String foreachContent) {
    for(Element drawingElement:wDrawingList) {
      Element docPr = (Element) drawingElement.selectSingleNode("descendant::wp:docPr");
      String domainName = docPr.attributeValue("descr");
      if(domainName == null) throw new RuntimeException("图片没有填写”替换文字“，请检查");
      // 如果时占位图片或者固定图片，不在这里处理
      if(domainName.matches("^placeholder.*?|^static.*?")) continue;
      // 获取foreach中item的内容
      String foreachItemContent = domainName;
      addPictureElement(document,drawingElement, foreachContent, foreachItemContent);
    }
  }
  private static void addPictureElement(Document document, Element drawingElement, String foreachContent, String foreachItemContent) {
    // 生成唯一id，用于关联w:drawing中rId---Relationship中Target---pkg:part中的binaryData
    String uuid = UUID.randomUUID().toString().replace("-", "");
    Element blip = (Element) drawingElement.selectSingleNode("descendant::*[namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main' and local-name()='blip']");
    blip.addAttribute("embed", "rId_" + uuid + "_$!{foreach.index}");
    // 新增Relationship循环开始标签
    Element relationshipForeachBegin = DocumentHelper.createElement(TEMP_TAG);
    relationshipForeachBegin.setText(foreachContent);
    // 新增Relationship循环内容标签
    Element relationshipForeachContent = DocumentHelper.createElement(QName.get("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"));
    relationshipForeachContent.addAttribute("Id", "rId_" + uuid + "_$!{foreach.index}");
    relationshipForeachContent.addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    relationshipForeachContent.addAttribute("Target", "media/image_" + uuid + "_$!{foreach.index}.png");
    // 新增Relationship循环结束标签
    Element relationshipForeachEnd = DocumentHelper.createElement(TEMP_TAG);
    relationshipForeachEnd.setText("#end");
    // 将Relationship循环标签添加到Relationships中
    Element pkgPart = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.microsoft.com/office/2006/xmlPackage' and local-name()='part' and @pkg:name='/word/_rels/document.xml.rels']");
    Element relationships = (Element) pkgPart.selectSingleNode("descendant::*[namespace-uri()='http://schemas.openxmlformats.org/package/2006/relationships' and local-name()='Relationships']");
    relationships.elements().add(relationshipForeachBegin);
    relationships.elements().add(relationshipForeachContent);
    relationships.elements().add(relationshipForeachEnd);
    // 新增pkg:part标签开始标签
    Element pkgPartForeachBegin = DocumentHelper.createElement(TEMP_TAG);
    pkgPartForeachBegin.setText(foreachContent);
    // 新增pkg:part循环内容标签
    Element pkgPartForeachContent = DocumentHelper.createElement("pkg:part");
    pkgPartForeachContent.addAttribute("pkg:name", "/word/media/image_" + uuid + "_$!{foreach.index}.png");
    pkgPartForeachContent.addAttribute("pkg:contentType", "image/png");
    pkgPartForeachContent.addAttribute("pkg:compression", "store");
    Element binaryData = DocumentHelper.createElement("pkg:binaryData");
    binaryData.setText(foreachItemContent);
    pkgPartForeachContent.add(binaryData);
    // 新增pkg:part循环结束标签
    Element pkgPartForeachEnd = DocumentHelper.createElement(TEMP_TAG);
    pkgPartForeachEnd.setText("#end");
    // 将pkg:part循环标签添加到pkg:package中
    Element pkgPackage = (Element) document.selectSingleNode("//*[namespace-uri()='http://schemas.microsoft.com/office/2006/xmlPackage' and local-name()='package']");
    pkgPackage.elements().add(pkgPartForeachBegin);
    pkgPackage.elements().add(pkgPartForeachContent);
    pkgPackage.elements().add(pkgPartForeachEnd);
  }

  /**
   * picture元素都处理好了，之后调用这个方法，将多余的属性去掉
   * 不能提前调用，可能把需要的属性干掉，其他处理图片的方法无法使用
   * 将前面 wp:docPr 的descr属性删除，将pic:cNvPr 的descr属性删除
   * @param document 文档元素
   */
  public static void removeRedundantPictureAttribute(Document document) {
    List<Element> docPrList = document.selectNodes("//*[local-name()='docPr' and contains(@descr, '$!{')  and contains(@descr, '}')]");
    for(Element element:docPrList) {
      Element inline = element.getParent();
      // 将 wp:docPr 的 descr 属性删除
      Attribute docPrDescr = element.attribute("descr");
      docPrDescr.detach();
      // 将 pic:cNvPr 的 descr 属性删除
      Element piccNvPr = (Element)inline.selectSingleNode("descendant::*[namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/picture' and local-name()='cNvPr']");
      Attribute piccNvPrDescr = piccNvPr.attribute("descr");
      if(piccNvPrDescr!=null) piccNvPrDescr.detach();
    }
  }
}