package com.example.report.util;

import com.spire.doc.FileFormat;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.RuntimeConstants;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;
import org.apache.velocity.tools.ToolContext;
import org.apache.velocity.tools.ToolManager;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Properties;

/**
 * 从xml生成doc，测试java将数据填到具体的对象中，使用velocity模板进行填充
 */
public class VelocityToolsTest {
  public static String TEST_PATH = "C:\\Users\\admin\\Desktop\\test\\";
  public static void main(String[] args) throws IOException {
    System.out.println("hello, world, 你好！");
    test();
  }

  public static void test() throws IOException {
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
    String name = dateFormat.format(new Date());
//    1 创建引擎对象
    VelocityEngine engine = new VelocityEngine();
//        2 设置模板的加载路径，设置了两个加载路径
    engine.setProperty(Velocity.RESOURCE_LOADERS,"class,file");
    engine.setProperty(Velocity.RESOURCE_LOADER_CLASS, ClasspathResourceLoader.class.getName());
    engine.setProperty(Velocity.FILE_RESOURCE_LOADER_PATH, TEST_PATH);
//        3  初始化引擎
    engine.init();
//        4  加载tools.xml配置文件，配置文件是配置在项目的resources中的
    ToolManager toolManager = new ToolManager();
    toolManager.configure("vm/tools.xml");
//        5  加载模板，模板是配置在文件夹中的
    Template template = engine.getTemplate("vm/tools.vm");
//        6  设置数据
    ToolContext context = toolManager.createContext();
    context.put("now",new Date());
//        7 合并数据到模板
    FileWriter fileWriter = new FileWriter("C:\\Users\\admin\\Desktop\\test\\" + name + ".txt");
    template.merge(context,fileWriter);
//        8 释放资源
    fileWriter.close();
  }
}

