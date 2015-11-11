package com.maplemart.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

/**
 * 用来导出word文件工具类
 * @author HuangYongFeng
 */
public class WordWebUtil
{
	private static Configuration configuration = null;
	private static Template template = null;
	static
	{
		try
		{
			configuration = new Configuration();
			configuration.setDefaultEncoding("utf-8");
			// 加载模板文件(到工程里的包路径,不用到文件名)
			configuration.setClassForTemplateLoading(WordWebUtil.class, "/config/templates");
			// 设置对象包装器
			configuration.setObjectWrapper(new DefaultObjectWrapper());
			// 设置异常处理器
			configuration.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
			// 定义Template对象,注意/config/templates包下的模板类型名字与fileName要一致
			Properties pro = new Properties();
			InputStream is = WordWebUtil.class.getResourceAsStream("/templates.properties");
			pro.load(is);
			template = configuration.getTemplate(pro.getProperty("fileName"));// wordTemplate.xml
		}
		catch (Exception e)
		{
			e.printStackTrace();
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据Doc模板生成word文件
	 * @author 黄永丰
	 * @createtime 2015年11月11日
	 * @param dataMap 需要填入模板的数据
	 * @param fileName 导出文件名称
	 */
	public static File createDoc(Map<String, Object> dataMap,String fileName)
	{
		try
		{
			File outFile = new File(fileName);
			FileOutputStream fos = new FileOutputStream(outFile);
			// 这个地方不能使用FileWriter因为需要指定编码类型否则生成的Word文档会因为有无法识别的编码而无法打开
			OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
			Writer out = new BufferedWriter(osw);
			template.process(dataMap,out);
			out.close();
			fos.close();
			return outFile;
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		return null;
	}

	/** 测试 */
	public static void main(String[] args)
	{
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("titleName", 1);
		dataMap.put("userName", 2);
		dataMap.put("orgName", 3);
		dataMap.put("deptName", 4);
		WordWebUtil.createDoc(dataMap,"E:\\outFile.doc");
	}
}
