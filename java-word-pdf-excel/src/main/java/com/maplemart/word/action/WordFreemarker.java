package com.maplemart.word.action;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.maplemart.word.util.WordWebUtil;


public class WordFreemarker extends HttpServlet
{

	private static final long serialVersionUID = -7450583995510823452L;

	@Override
	protected void service(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException
	{
		Map<String, Object> dataMap = new HashMap<String, Object>();
		dataMap.put("titleName", null);
		dataMap.put("userName", 2);
		dataMap.put("orgName", 3);
		dataMap.put("deptName", 4);
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/x-msdownload");
		response.addHeader("Content-Disposition", "attachment;filename=model.doc");
		try
		{
			WordWebUtil.createDoc(dataMap,response.getWriter());
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

}
