package com.example.report.util;

import com.spire.doc.FileFormat;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.RuntimeConstants;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

/**
 * 从xml生成doc，测试java将数据填到具体的对象中，使用velocity模板进行填充
 */
public class VelocityTest {
  public static String TEST_PATH = "C:\\Users\\admin\\Desktop\\test\\";
  public static String DOCX_FILE = "ahello.docx";
  public static String XML_FILE = "ahello.xml";
  public static String TARGET_FILE = "ahello.doc";
  public static void main(String[] args) {
    System.out.println("hello, world, 你好！");
    test();
  }

  public static void generateXML() {
    String templatePath = TEST_PATH + DOCX_FILE;
    //注册Spire.Doc
    com.spire.license.LicenseProvider.setLicenseKey("HmFcwoMBAI0cOTiuZZLGencD92lphsezQhdIkPgJk8Hmjq/OJx3XMSXgz81ybjFm0Tm5aoJ+Oh59lNKd7Wr4XlgeAnW8uNkz8ZMuhQEUXGMtGUv59qr5W9LA1ZUDBPbOFuGDPBhTTW00PaQu21qpSTaSQ0lm6Zp5PnhwOTwufKagtVSLfBzD/OYfVVIsoBFC2XPR0Qj55o2RcJXrF82DvnFFpdrcAemnAUEwzWw9qYEVfLfwTbDKXmngIs673fdwFxJM/u1fQhZcEGlcN3veup1ITr8hqo4CskJftZY9Ewt7Hf+omDzs1oKGNno4XM6Smm45LKVkv0iKj4Ms1x2eyEVolmZCmqtagzwS+ICbZAja0hBSbWwCOMk6qXa5gcNrMxRXU/TOxKJzb2FcDBnB5Z1mqWCUNOlWtjh4Yp+PR9L9Wb24D9rrwS9kFmPDTvhHlAxppKi5bE8RHecJMia9xL6K5zI2/srGCOGan+M+pPX0EZ5bpaeFutE0uNvciqxCXhHtz4apzk9d//ka48opuwngtqYk6k743sApf2UD7WH5jg3HtxoH8XWnFm/dN/Zq6EM+b4YHmGG1nW7ulLbbrMwQhDNhdOPRcVws+ioy6jU7LlvKDDweTZMInzbn+ZDZ8SrCLXCNNiJ/zD84AEGxXCyXcbCDOB9Ru4X92Hpg+0+VicyCshgbnJRwX0/QT41ikZZCvnR1eZUAPSn31MSPwVKsv34NwGQkMAFCPKa6iSsrr9hNrgmHplY9eWo0xWUuihXiIzceCT3JhkOlODsxK2NgcMoTYW0Lw9yCkhP0FaKeFWlOunu4uKR1R9tsUa3Jf6trzqk4XnPDQoo+xf/R2ggwwpcAaSn15cO4EtzWOpvYaD9UdOeKlTve1heUkEDuBv+lDyJa8VPrzfSIPyzL2s/pw8BhpqC0s60ovKDFLOL37nk3xRv6CcoaPZEY/Jo5LPJyyDpZVNpjMDmrU72gu+O2ulnPkfJ8s+oi3V7IMAfFFFyatGb252p1RlADjM8QxY+FhF2vUoTY1HSxJYvbSDUYIT6iN9Dgvv6ZAyujPbfaUnxLA+e0AabfcI2dFKKW25b3aqxug6ClK6ZgrKwnfs8v+ihHRE6PwL6yVUThLHmaumeagGqu5EZkOMw76Kfx6fPSIMBJEpXaTNPV6TFaatb+IQHusrmcY4sVEEaPB9QPg0tbUGwHyirPfAMmGJzMjyMPmJCCTb6JObwuBsPORMKBlohxY6OP5ANboHRspKozwO7XDCM9Kvqr3uMQgoWuIu5/JR6KWw8viF9S4j62oDTKq4ida1+Nr8kIPY9zvIWnWAuA/gplCutQnjOlXZNqkwr7DeeNcPpseX/9YJCtmI/Fm4+07NhccFYivnzYMG8OjQIdQwLxgpAOIiBh4rw5Ks1pcKD+7RWXlVuFzCj9zbmU6OrqF10irb9PRwwStK+ndOYrTwMaCYDZ397Hj4LufO3VhfiUh+G8o3y/U7h5WJkaAh10fMfocNsybR5zPyZDe8HOIeCdwMNXwmJYRgDns6oxOEN8oQzxvE4ihUqRxA==");
    //加载模板
    com.spire.doc.Document document = new com.spire.doc.Document();
    document.loadFromFile(templatePath, FileFormat.Docx);
    //将模板另存为xml
    document.saveToFile(TEST_PATH + XML_FILE, FileFormat.Word_Xml);
    System.out.println("word转xml成功");
  }
  public static void test() {
    Properties p = new Properties();
    // p.setProperty(Velocity.FILE_RESOURCE_LOADER_PATH, "");
    p.setProperty(Velocity.ENCODING_DEFAULT, "UTF-8");
    p.setProperty(Velocity.INPUT_ENCODING, "UTF-8");
    p.setProperty(Velocity.OUTPUT_ENCODING, "UTF-8");
    // 初始化模板引擎
    VelocityEngine ve = new VelocityEngine();
    // 类路径
//    ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "classpath");
//    ve.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
    // 绝对路径
    ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "file");
    ve.setProperty(RuntimeConstants.FILE_RESOURCE_LOADER_PATH, TEST_PATH);
    ve.init();
    // 获取模板文件
    Template t = ve.getTemplate(XML_FILE);
    // 设置变量
    VelocityContext ctx = new VelocityContext();
    ctx.put("greet", "Velocity");
    PatientInfo pi = new PatientInfo("张三", "13512341234", "男", 16);
    ctx.put("pi", pi);
    List<PatientInfo> list = new ArrayList<>();
    list.add(new PatientInfo("lisi1", "13512345678", "male", 18));
    list.add(new PatientInfo("lisi2", "13512345678", "male", 18));
    list.add(new PatientInfo("lisi3", "13512345678", "male", 18));
    ctx.put("list", list);
    ctx.put("picBase64", PictureData.getPicString());
    // 输出
//    StringWriter sw = new StringWriter();
//    t.merge(ctx, sw);
//    System.out.println(sw.toString());
    File file = new File(TEST_PATH + TARGET_FILE);
    BufferedWriter bw = null;
    try {
      bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8));
      t.merge(ctx, bw);
      bw.flush();;
      bw.close();
    } catch(Exception e) {
      e.printStackTrace();
    }
  }
}

