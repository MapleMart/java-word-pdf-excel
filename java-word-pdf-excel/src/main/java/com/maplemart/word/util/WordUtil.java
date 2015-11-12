package com.maplemart.word.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

/**
 * 用来导出word文件工具类
 * @author HuangYongFeng
 */
public class WordUtil
{
	private Configuration configure = null;
	public WordUtil()
	{
		configure = new Configuration();
		configure.setDefaultEncoding("utf-8");
	}

	/**
	 * 根据Doc模板生成word文件
	 * @author 黄永丰
	 * @createtime 2015年11月11日
	 * @param dataMap 需要填入模板的数据
	 * @param fileName 模板文件名称
	 * @param savePath 保存路径
	 */
	public void createDoc(Map<String, Object> dataMap, String fileName, String savePath)
	{
		try
		{
			// 加载模板文件(到工程里的包路径,不用到文件名)
			configure.setClassForTemplateLoading(this.getClass(), "/config/templates");
			// 设置对象包装器
			configure.setObjectWrapper(new DefaultObjectWrapper());
			// 设置异常处理器
			configure.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
			// 定义Template对象,注意/config/templates包下的模板类型名字与fileName要一致
			Template template = configure.getTemplate(fileName + ".ftl");//wordTemplate.xml
			// 输出文档路径及名称
			File outFile = new File(savePath);
			FileOutputStream fos = new FileOutputStream(outFile);
			OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
			Writer out = new BufferedWriter(osw);
			template.process(dataMap,out);
			out.close();
			fos.close();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
	
	/**测试*/
	public static void main(String[] args)
	{
		 Map<String, Object> dataMap = new HashMap<String, Object>();  
		 dataMap.put("titleName", 1);
		 dataMap.put("userName", 2);
		 dataMap.put("orgName", 3);
		 dataMap.put("deptName", 4);
		 new WordUtil().createDoc(dataMap, "wordTemplate", "E:\\outFile.doc");
	}
}





