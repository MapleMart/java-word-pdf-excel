<%@page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@page import="com.maplemart.util.*,java.util.*,java.io.*"%>
<%
	// 提示：在调用工具类生成Word文档之前应当检查所有字段是否完整  
    // 否则Freemarker的模板殷勤在处理时可能会因为找不到值而报错 这里暂时忽略这个步骤了  
    File file = null;  
    InputStream is = null;  
    ServletOutputStream sos = null;  
    try {  
    	String sheetTitle = new String("outFile.doc".getBytes(), "ISO-8859-1");
    	Map<String, Object> dataMap = new HashMap<String, Object>();  
    	dataMap.put("titleName", 1);
    	dataMap.put("userName", 2);
    	dataMap.put("orgName", 3);
    	dataMap.put("deptName", 4);
        // 调用工具类WordWebUtil的createDoc方法生成Word文档  
        file = WordWebUtil.createDoc(dataMap,"E:\\outFile.doc");
        is = new FileInputStream(file);  
          
        response.setCharacterEncoding("utf-8");  
        response.setContentType("application/msword");  
        // 设置浏览器以下载的方式处理该文件默认名为resume.doc  
        response.addHeader("Content-Disposition", "attachment;filename="+sheetTitle);  
        
        sos = response.getOutputStream();  
        byte[] buffer = new byte[1024];  // 缓冲区  
        int bytesToRead = -1;  
        // 通过循环将读入的Word文件的内容输出到浏览器中  
        while((bytesToRead = is.read(buffer)) != -1) {  
        	sos.write(buffer, 0, bytesToRead);  
        }  
    } finally {  
        if(is != null) is.close();  
        if(out != null) out.close();  
        if(file != null) file.delete(); // 删除临时文件  
    }  
%>

