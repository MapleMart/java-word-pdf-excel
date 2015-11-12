package com.maplemart.excel.util;

import java.io.InputStream;
import java.io.Writer;
import java.util.Map;
import java.util.Properties;

import com.maplemart.word.util.WordWebUtil;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

/**
 * 用来导出excel工具类
 * @author 黄永丰
 * @createtime 2015年11月12日
 * @version 1.0
 */
public class ExcelUtil
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
			configuration.setClassForTemplateLoading(WordWebUtil.class, "/templates/excel");
			// 设置对象包装器
			configuration.setObjectWrapper(new DefaultObjectWrapper());
			// 设置异常处理器
			configuration.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
			// 定义Template对象,注意/config/templates包下的模板类型名字与fileName要一致
			Properties pro = new Properties();
			InputStream is = WordWebUtil.class.getResourceAsStream("/templates/templates.properties");
			pro.load(is);
			template = configuration.getTemplate(pro.getProperty("excel.template.fileName"));// wordTemplate.xml
		}
		catch (Exception e)
		{
			e.printStackTrace();
			throw new RuntimeException(e);
		}
	}
	
	/**
	 * 根据Excel模板生成Excel文件
	 * @author 黄永丰
	 * @createtime 2015年11月11日
	 * @param dataMap 需要填入模板的数据
	 * @param out 转出流
	 */
	public static void createExcel(Map<String, Object> dataMap,Writer out)
	{
		try
		{
			template.process(dataMap,out);
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
