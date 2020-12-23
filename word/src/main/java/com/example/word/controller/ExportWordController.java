package com.example.word.controller;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.util.ClassUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.example.word.common.MoneyUtils;
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
	

	/**
	 * 销售订单信息导出word --- poi-tl（包含动态表格）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord3")
	public void exportDataWord3(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
	
	        // TODO 渲染其他类型的数据请参考官方文档
	        DecimalFormat df = new DecimalFormat("######0.00");   
	        Calendar now = Calendar.getInstance(); 
	        double money = 0;//总金额
	        //组装表格列表数据
	        List<Map<String,Object>> detailList=new ArrayList<Map<String,Object>>();
	        for (int i = 0; i < 6; i++) {
	        	 Map<String,Object> detailMap = new HashMap<String, Object>();
	        	 detailMap.put("index", i+1);//序号
	        	 detailMap.put("title", "商品"+i);//商品名称
	        	 detailMap.put("product_description", "套");//商品规格
	        	 detailMap.put("buy_num", 3+i);//销售数量
	        	 detailMap.put("saleprice", 100+i);//销售价格
	        	
	        	 double saleprice=Double.valueOf(String.valueOf(100+i));
	             Integer buy_num=Integer.valueOf(String.valueOf(3+i));
	             String buy_price=df.format(saleprice*buy_num);
	             detailMap.put("buy_price", buy_price);//单个商品总价格
	             money=money+Double.valueOf(buy_price);
	             
	             detailList.add(detailMap);
	        }
	        //总金额
	        String order_money=String.valueOf(money);
	        //金额中文大写
	        String money_total = MoneyUtils.change(money);
	      	
	      	String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
	      	String resource =basePath+"order1.docx";//word模板地址
	        //渲染表格
	        HackLoopTableRenderPolicy  policy = new HackLoopTableRenderPolicy();
	        Configure config = Configure.newBuilder().bind("detailList", policy).build();
	        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(
	        		new HashMap<String, Object>() {{
			        put("detailList", detailList);
			        put("order_number", "2356346346645");
			        put("y", now.get(Calendar.YEAR));//当前年
			        put("m", (now.get(Calendar.MONTH) + 1));//当前月
			        put("d", now.get(Calendar.DAY_OF_MONTH));//当前日
			        put("order_money",order_money);//总金额
			        put("money_total",money_total);//金额中文大写
			    }}
			);
	        //=================生成文件保存在本地D盘某目录下=================
	        String temDir="D:/mimi/"+File.separator+"file/word/"; ;//生成临时文件存放地址
			//生成文件名
			Long time = new Date().getTime();
	        // 生成的word格式
	        String formatSuffix = ".docx";
	        // 拼接后的文件名
	        String fileName = time + formatSuffix;//文件名  带后缀
	        
	        FileOutputStream fos = new FileOutputStream(temDir+fileName);
	        template.write(fos);
	        //=================生成word到设置浏览默认下载地址=================
            // 设置强制下载不打开
            response.setContentType("application/force-download");
            // 设置文件名
            response.addHeader("Content-Disposition", "attachment;fileName=" + fileName);
            OutputStream out = response.getOutputStream();
            template.write(out);
            out.flush();
            out.close();
            template.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

	}
}
