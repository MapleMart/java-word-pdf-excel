package com.maplemart.excel.action;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.maplemart.excel.util.ExcelUtil;


public class ExcelFreemarker extends HttpServlet
{

	private static final long serialVersionUID = -7450583995510823452L;

	@Override
	protected void service(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
	{
		//测试Excel文件生成
		//内容
		List<Map<String,Object>> list = new ArrayList<Map<String,Object>>(); 
		for (int i = 0; i < 5; i++)
		{
			Map<String, Object> dataMap = new HashMap<String, Object>();
			dataMap.put("username", 1+i);//用来和Excel里的${username}里的名称要一样
			dataMap.put("password", 2+i);
			dataMap.put("name", 3+i);
			dataMap.put("age", 4+i);
			dataMap.put("address", 5+i);
			list.add(dataMap);
		}
		//标题和存放内容
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("list", list);
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/x-msdownload");
		response.addHeader("Content-Disposition", "attachment;filename=model.xls");
		try
		{
			ExcelUtil.createExcel(dataMap,response.getWriter());
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
