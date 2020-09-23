package com.example.word.controller;


import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;
import com.deepoove.poi.data.PictureRenderData;
import com.example.word.common.WordUtil;
import com.example.word.common.WordUtil2;

import cn.afterturn.easypoi.word.entity.WordImageEntity;

/**
 * 导出Word
 * @author Administrator
 *
 */
@RequestMapping("/auth/exportWord/")
@RestController
public class ExportWordController {

	/**
	 * 导出word首页
	 */
	@RequestMapping(value ="/index")
	public ModelAndView toIndex(HttpServletRequest request){
		return new ModelAndView("export/index");
	}
	/**
	 * 用户信息导出word --- poi-tl
	 * @throws IOException 
	 */
	@RequestMapping("/exportUserWord")
	public void exportUserWord(HttpServletRequest request,HttpServletResponse response) throws IOException{
//		response.setContentType("application/vnd.ms-excel");  
		response.setContentType("application/octet-stream");  
        response.setHeader("Content-Disposition", "attachment;filename=" +new SimpleDateFormat("yyyyMMddHHmm").format(new Date())+".docx"); 
		Map<String, Object> params = new HashMap<>();
        // 渲染文本
        params.put("name", "张三");
        params.put("position", "开发工程师");
        params.put("entry_time", "2020-07-30");
        params.put("province", "江苏省");
        params.put("city", "南京市");
        // 渲染图片
        params.put("picture", new PictureRenderData(120, 120, "D:\\cssTest\\square.png"));
        // TODO 渲染其他类型的数据请参考官方文档
        
        //word模板地址放在src/main/webapp/下
		//表示到项目的根目录（webapp）下，要是想到目录下的子文件夹，修改"/"即可
		String path = request.getSession().getServletContext().getRealPath("/");
        String templatePath = path+"template/user.docx";
        
		WordUtil.downloadWord(response.getOutputStream(), templatePath, params);
	}
	
	/**
	 * 用户信息导出word  --- easypoi
	 * @throws IOException 
	 */
	@RequestMapping("/exportUserWord2")
	public void exportUserWord2(HttpServletRequest request,HttpServletResponse response) throws IOException{
		Map<String, Object> params = new HashMap<>();
		
		//表示到项目的根目录（webapp）下，要是想到目录下的子文件夹，修改"/"即可
		String path = request.getSession().getServletContext().getRealPath("/");
		
		//word模板地址放在src/main/webapp/下
		String templatePath =path+"template/user2.docx";//word模板地址
		
		// 渲染文本
		params.put("name", "张三");
		params.put("position", "开发工程师");
		params.put("entry_time", "2020-07-30");
		params.put("province", "江苏省");
		params.put("city", "南京市");
		
		// 渲染图片
		WordImageEntity image = new WordImageEntity();
        image.setHeight(120);
        image.setWidth(120);
        image.setUrl("D:\\cssTest\\square.png");
        image.setType(WordImageEntity.URL);
        params.put("picture", image);
		// TODO 渲染其他类型的数据请参考官方文档
        
		
		String temDir="D:/mimi/"+File.separator+"file/word/"; ;//生成临时文件存放地址
		
		//生成文件名
		Long time = new Date().getTime();
        // 生成的word格式
        String formatSuffix = ".docx";
        // 拼接后的文件名
        String fileName = time + formatSuffix;//文件名  带后缀
        //导出word
		WordUtil2.exportWord(templatePath, temDir, fileName, params, request, response);
	}
}
