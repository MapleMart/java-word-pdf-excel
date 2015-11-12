package com.maplemart.word.util;

import java.io.InputStream;
import java.io.Writer;
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
			configuration.setClassForTemplateLoading(WordWebUtil.class, "/templates/word");
			// 设置对象包装器
			configuration.setObjectWrapper(new DefaultObjectWrapper());
			// 设置异常处理器
			configuration.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
			// 定义Template对象,注意/config/templates包下的模板类型名字与fileName要一致
			Properties pro = new Properties();
			InputStream is = WordWebUtil.class.getResourceAsStream("/templates/templates.properties");
			pro.load(is);
			template = configuration.getTemplate(pro.getProperty("word.template.fileName"));// wordTemplate.xml
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
	 * @param out 转出流
	 */
	public static void createDoc(Map<String, Object> dataMap,Writer out)
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
