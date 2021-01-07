package com.example.word.controller;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.springframework.util.ClassUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.policy.HackLoopTableRenderPolicy;
import com.example.word.common.MoneyUtils;
import com.example.word.common.WordUtil;
import com.example.word.common.WordUtil2;
import com.example.word.common.mergeCell.DetailData;
import com.example.word.common.mergeCell.DetailTablePolicy;
import com.example.word.common.mergeCell.PaymentData;
import com.example.word.common.mergeCell2.DetailData2;
import com.example.word.common.mergeCell2.DetailTablePolicy2;
import com.example.word.common.mergeCell2.PaymentData2;
import com.example.word.common.mergeCell3.DetailData3;
import com.example.word.common.mergeCell3.DetailTablePolicy3;
import com.example.word.common.mergeCell3.DetailTablePolicy4;
import com.example.word.common.mergeCell3.PaymentData3;

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
	 * 销售订单信息导出word --- poi-tl（包含动态行表格）
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（包含两个动态行表格）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWordD4")
	public void exportDataWordD4(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			DecimalFormat df = new DecimalFormat("######0.00");   
			Calendar now = Calendar.getInstance(); 
			double money = 0;//总金额
			//组装表格列表数据
			List<Map<String,Object>> typeList=new ArrayList<Map<String,Object>>();
			for (int i = 0; i < 2; i++) {
				Map<String,Object> detailMap = new HashMap<String, Object>();
				detailMap.put("index", i+1);//序号
				if(i == 0){
					detailMap.put("sub_type", "监督技术装备");//商品所属大类名称
				}else if(i == 1){
					detailMap.put("sub_type", "火灾调查装备");//商品所属大类名称
				}else if(i == 2){
					detailMap.put("sub_type", "工程验收装备");//商品所属大类名称
				}
				
				double saleprice=Double.valueOf(String.valueOf(100+i));
				Integer buy_num=Integer.valueOf(String.valueOf(3+i));
				String buy_price=df.format(saleprice*buy_num);
				detailMap.put("buy_price", buy_price);//所属大类总价格
				money=money+Double.valueOf(buy_price);
				typeList.add(detailMap);
			}
			//组装表格列表数据
			List<Map<String,Object>> detailList=new ArrayList<Map<String,Object>>();
			for (int i = 0; i < 3; i++) {
				Map<String,Object> detailMap = new HashMap<String, Object>();
				detailMap.put("index", i+1);//序号
				if(i == 0 || i == 1){
					detailMap.put("product_type", "二级分类1");//商品二级分类
				}else{
					detailMap.put("product_type", "二级分类2");//商品二级分类
				}
				detailMap.put("title", "商品"+i);//商品名称
				detailMap.put("product_description", "套");//商品规格
				detailMap.put("buy_num", 3+i);//销售数量
				detailMap.put("saleprice", 100+i);//销售价格
				detailMap.put("technical_parameter", "技术参数"+i);//技术参数
				detailList.add(detailMap);
			}
			
			//总金额
			String order_money=String.valueOf(money);
			//金额中文大写
			String money_total = MoneyUtils.change(money);
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"orderD2.docx";//word模板地址
			//渲染表格  动态行
			HackLoopTableRenderPolicy  policy = new HackLoopTableRenderPolicy();
			Configure config = Configure.newBuilder()
					.bind("typeList", policy).bind("detailList", policy).build();
			
			XWPFTemplate template = XWPFTemplate.compile(resource, config).render(
					new HashMap<String, Object>() {{
						put("typeList", typeList);
						put("detailList",detailList);
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（包含动态行表格、循环列表中的动态行表格）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord4")
	public void exportDataWord4(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			DecimalFormat df = new DecimalFormat("######0.00");   
			Calendar now = Calendar.getInstance(); 
			double money = 0;//总金额
			//组装表格列表数据
			List<Map<String,Object>> typeList=new ArrayList<Map<String,Object>>();
			for (int i = 0; i < 2; i++) {
				Map<String,Object> detailMap = new HashMap<String, Object>();
				detailMap.put("index", i+1);//序号
				if(i == 0){
					detailMap.put("sub_type", "监督技术装备");//商品所属大类名称
				}else if(i == 1){
					detailMap.put("sub_type", "火灾调查装备");//商品所属大类名称
				}else if(i == 2){
					detailMap.put("sub_type", "工程验收装备");//商品所属大类名称
				}
				
				double saleprice=Double.valueOf(String.valueOf(100+i));
				Integer buy_num=Integer.valueOf(String.valueOf(3+i));
				String buy_price=df.format(saleprice*buy_num);
				detailMap.put("buy_price", buy_price);//所属大类总价格
				money=money+Double.valueOf(buy_price);
				typeList.add(detailMap);
			}
			//组装表格列表数据
			List<Map<String,Object>> detailList=new ArrayList<Map<String,Object>>();
			for (int i = 0; i < 3; i++) {
				Map<String,Object> detailMap = new HashMap<String, Object>();
				detailMap.put("index", i+1);//序号
				if(i == 0 || i == 1){
					detailMap.put("product_type", "二级分类1");//商品二级分类
				}else{
					detailMap.put("product_type", "二级分类2");//商品二级分类
				}
				detailMap.put("title", "商品"+i);//商品名称
				detailMap.put("product_description", "套");//商品规格
				detailMap.put("buy_num", 3+i);//销售数量
				detailMap.put("saleprice", 100+i);//销售价格
				detailMap.put("technical_parameter", "技术参数"+i);//技术参数
				detailList.add(detailMap);
			}
			
			
			List<Map<String,Object>> tList=new ArrayList<Map<String,Object>>();
			Map<String,Object> tMap = new HashMap<String, Object>();
			tMap.put("index", 1);
			tMap.put("sub_type", "监督技术装备");
			tMap.put("detailList", detailList);
			tMap.put("buy_price", 100);
			tList.add(tMap);
			
			tMap = new HashMap<String, Object>();
			tMap.put("index", 2);
			tMap.put("sub_type", "火灾调查装备");
			tMap.put("detailList", detailList);
			tMap.put("buy_price", 200);
			tList.add(tMap);
			
			tMap = new HashMap<String, Object>();
			tMap.put("index", 3);
			tMap.put("sub_type", "工程验收装备");
			tMap.put("detailList", detailList);
			tMap.put("buy_price", 300);
			tList.add(tMap);
			
			
			//总金额
			String order_money=String.valueOf(money);
			//金额中文大写
			String money_total = MoneyUtils.change(money);
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"order2.docx";//word模板地址
			//渲染表格  动态行
			HackLoopTableRenderPolicy  policy = new HackLoopTableRenderPolicy();
			Configure config = Configure.newBuilder()
					.bind("typeList", policy).bind("detailList", policy).build();
			
			XWPFTemplate template = XWPFTemplate.compile(resource, config).render(
					new HashMap<String, Object>() {{
						put("typeList", typeList);
						put("typeProducts",tList);
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（合并单元格----货物明细\人工费）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord5")
	public void exportDataWord5(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"order3.docx";//word模板地址
			
			PaymentData datas = new PaymentData();
			TableStyle rowStyle = new TableStyle();
			rowStyle = new TableStyle();
		    rowStyle.setAlign(STJc.CENTER);
			DetailData detailTable = new DetailData();
	        RowRenderData good = RowRenderData.build("4", "墙纸", "书房+卧室", "1500", "/", "400", "1600");
	        good.setRowStyle(rowStyle);
	        List<RowRenderData> goods = Arrays.asList(good, good, good);
	        RowRenderData labor = RowRenderData.build("油漆工", "2", "200", "400");
	        labor.setRowStyle(rowStyle);
	        List<RowRenderData> labors = Arrays.asList(labor, labor, labor, labor);
	        detailTable.setGoods(goods);
	        detailTable.setLabors(labors);
	        datas.setDetailTable(detailTable);
	        datas.setTotal("1000");
	        
			Configure config = Configure.newBuilder().bind("detail_table", new DetailTablePolicy()).build();
			
	        XWPFTemplate template = XWPFTemplate.compile(resource, config).render(datas);
	        
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（合并单元格（一个列表下的合并行）--商品订单明细）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord6")
	public void exportDataWord6(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"order4.docx";//word模板地址
			
			PaymentData2 datas = new PaymentData2();
			TableStyle rowStyle = new TableStyle();
			rowStyle = new TableStyle();
			rowStyle.setAlign(STJc.CENTER);
			DetailData2 detailTable = new DetailData2();
			List<RowRenderData> plists =new ArrayList<RowRenderData>();
		    for(int i=0;i<6;i++){
		    	String typeName="二级分类1";
		    	if(i == 3 || i == 4){
		    		typeName="二级分类2";
		    	}else if(i == 5){
		    		typeName="二级分类3";
		    	}
		    	String index = String.valueOf(i+1);
		    	RowRenderData plist = RowRenderData.build(index, typeName, "商品"+i, "套", "2","1000.00","技术参数"+i);
		    	plist.setRowStyle(rowStyle);
				plists.add(plist);
				
		    }
		    //二级分类 分组统计   商品个数
		    List<Map<String,Object>> tlists = new ArrayList<Map<String,Object>>();
		    Map<String,Object> map = new HashMap<String, Object>();
		    map.put("typeName", "二级分类1");
		    map.put("listSize", "3");
		    tlists.add(map);
		    map = new HashMap<String, Object>();
		    map.put("typeName", "二级分类2");
		    map.put("listSize", "2");
		    tlists.add(map);
		    map = new HashMap<String, Object>();
		    map.put("typeName", "二级分类3");
		    map.put("listSize", "1");
		    tlists.add(map);
		    
		    
			detailTable.setPlists(plists);
			detailTable.setTlists(tlists);
			datas.setDetailTable(detailTable);
			datas.setTotal("1000");
			
			Configure config = Configure.newBuilder().bind("detail_table", new DetailTablePolicy2()).build();
			
			XWPFTemplate template = XWPFTemplate.compile(resource, config).render(datas);
			
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（合并单元格（循环列表下的合并行）--商品订单明细）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord7")
	public void exportDataWord7(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"order5.docx";//word模板地址
			
			PaymentData2 datas = new PaymentData2();
			TableStyle rowStyle = new TableStyle();
			rowStyle = new TableStyle();
			rowStyle.setAlign(STJc.CENTER);
			//组装循环体
			List<Map<String,Object>> typeLists = new ArrayList<Map<String,Object>>();
			for(int x=0;x<3;x++){
				DetailData2 detailTable = new DetailData2();
				List<RowRenderData> plists =new ArrayList<RowRenderData>();
				for(int i=0;i<6;i++){
					String typeName="二级分类1"+x;
					if(i == 3 || i == 4){
						typeName="二级分类2"+x;
					}else if(i == 5){
						typeName="二级分类3"+x;
					}
					String index = String.valueOf(i+1);
					RowRenderData plist = RowRenderData.build(index, typeName, "商品"+i, "套", "2","100","技术参数"+i);
					plist.setRowStyle(rowStyle);
					plists.add(plist);
					
				}
				//二级分类 分组统计   商品个数
				List<Map<String,Object>> tlists = new ArrayList<Map<String,Object>>();
				Map<String,Object> map = new HashMap<String, Object>();
				map.put("typeName", "二级分类1"+x);
				map.put("listSize", "3");
				tlists.add(map);
				map = new HashMap<String, Object>();
				map.put("typeName", "二级分类2"+x);
				map.put("listSize", "2");
				tlists.add(map);
				map = new HashMap<String, Object>();
				map.put("typeName", "二级分类3"+x);
				map.put("listSize", "1");
				tlists.add(map);
				
				
				detailTable.setPlists(plists);
				detailTable.setTlists(tlists);
				Map<String,Object> data= new HashMap<String, Object>();
				data.put("detail_table", detailTable);
				data.put("sub_type", "大类"+x);
				data.put("total_price", 100+x);
				typeLists.add(data);
			}
			datas.setTypeLists(typeLists);
			
			Configure config = Configure.newBuilder().bind("detail_table", new DetailTablePolicy2()).build();
			
			XWPFTemplate template = XWPFTemplate.compile(resource, config).render(datas);
			
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
	
	/**
	 * 销售订单信息导出word --- poi-tl（合并单元格（循环列表下的合并行）--商品订单明细、另加一个动态行表格）
	 * @throws IOException 
	 */
	@RequestMapping("/exportDataWord8")
	public void exportDataWord8(HttpServletRequest request,HttpServletResponse response) throws IOException{
		try {
			Map<String, Object> params = new HashMap<>();
			
			// TODO 渲染其他类型的数据请参考官方文档
			
			String basePath=ClassUtils.getDefaultClassLoader().getResource("").getPath()+"static/template/";
			String resource =basePath+"order6.docx";//word模板地址
			
			PaymentData3 datas = new PaymentData3();
			TableStyle rowStyle = new TableStyle();
			rowStyle = new TableStyle();
			rowStyle.setAlign(STJc.CENTER);
			//组装循环体
			List<Map<String,Object>> typeLists = new ArrayList<Map<String,Object>>();
			List<Map<String,Object>> typeTotalList = new ArrayList<Map<String,Object>>();
			for(int x=0;x<3;x++){
				DetailData3 detailTable = new DetailData3();
				List<RowRenderData> plists =new ArrayList<RowRenderData>();
				for(int i=0;i<6;i++){
					String typeName="二级分类1"+x;
					if(i == 3 || i == 4){
						typeName="二级分类2"+x;
					}else if(i == 5){
						typeName="二级分类3"+x;
					}
					String index = String.valueOf(i+1);
					RowRenderData plist = RowRenderData.build(index, typeName, "商品"+i, "套", "2","100","技术参数"+i);
					plist.setRowStyle(rowStyle);
					plists.add(plist);
					
				}
				//二级分类 分组统计   商品个数
				List<Map<String,Object>> tlists = new ArrayList<Map<String,Object>>();
				Map<String,Object> map = new HashMap<String, Object>();
				map.put("typeName", "二级分类1"+x);
				map.put("listSize", "3");
				tlists.add(map);
				map = new HashMap<String, Object>();
				map.put("typeName", "二级分类2"+x);
				map.put("listSize", "2");
				tlists.add(map);
				map = new HashMap<String, Object>();
				map.put("typeName", "二级分类3"+x);
				map.put("listSize", "1");
				tlists.add(map);
				
				
				detailTable.setPlists(plists);
				detailTable.setTlists(tlists);
				Map<String,Object> data= new HashMap<String, Object>();
				data.put("detail_table", detailTable);
				data.put("index", x+1);
				data.put("sub_type", "大类"+x);
				data.put("total_price", 100+x);
				typeLists.add(data);
			}
			datas.setTypeLists(typeLists);
			datas.setOrder_money("100");
			datas.setMoney_total("壹佰元整");
			
			Configure config = Configure.newBuilder().bind("detail_table", new DetailTablePolicy3())
					.bind("typeLists", new DetailTablePolicy4()).build();
			
			XWPFTemplate template = XWPFTemplate.compile(resource, config).render(datas);
			
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
