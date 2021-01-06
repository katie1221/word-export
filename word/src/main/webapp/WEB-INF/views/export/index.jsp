<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE html>
<html>
<head>
  <style>
    .m1{
      margin-left: 10px;
    }
    .s1{
      font-weight: bold;
      color:#1e9fff;
    }
    .m2{
      margin-bottom: 10px;
    }
  </style>
</head>
<body  class="easyui-layout">

    <h1 class="m1">SpringBoot导出word的几种方式：</h1>
    <div class="easyui-tabs" style="width:700px;height:250px;margin: 10px 0 0 10px;">
        <div title="方式一" style="padding:10px">
	        <div class="m2">方式一：使用<span class="s1">POI-tl</span>根据word模板动态生成word（包含文本、普通表格、图片）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord1();" data-options="iconCls:'icon-save'">导出word方式一</a>
        
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（包含动态行表格）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord3();" data-options="iconCls:'icon-save'">导出word（包含动态行表格）</a>
            
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（包含两个动态行表格）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWordD4();" data-options="iconCls:'icon-save'">导出word（包含两个动态行表格）</a>
            
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（包含动态行表格、循环列表下动态行表格）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord4();" data-options="iconCls:'icon-save'">导出word（包含动态行表格、循环列表下动态行表格）</a>
			
            
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（合并单元格1）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord5();" data-options="iconCls:'icon-save'">导出word（合并单元格（合并列）--货物明细\人工费）</a>
			
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（合并单元格2）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord6();" data-options="iconCls:'icon-save'">导出word（合并单元格（一个列表下的合并行）--商品订单明细）</a>
			<div style="font-size: 13px;color:#cccccc">如：合并第1列的第1行到第3行的单元格</div>
			
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（循环列表下的动态行表格以及合并单元格）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord7();" data-options="iconCls:'icon-save'">导出word（合并单元格（循环列表下的合并行）--商品订单明细）</a>
			
            <div class="m2" style="margin-top: 30px;">使用<span class="s1">POI-tl</span>根据word模板动态生成word（循环列表下的动态行表格以及合并单元格）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord8();" data-options="iconCls:'icon-save'">导出word（合并单元格（循环列表下的合并行）--商品订单明细、另加一个动态行表格）</a>
        </div>
         <div title="方式二" style="padding:10px">
	        <div class="m2">方式二：使用<span class="s1">easypoi</span>根据word模板动态生成word（包含文本、普通表格、图片）</div>
			<a href="#" class="easyui-linkbutton" onclick="doExportWord2();" data-options="iconCls:'icon-save'">导出word方式二</a>
        </div>
    </div>
	<script type="text/javascript">
	  //方式一导出word
	  function doExportWord1(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportUserWord";
	  }
	  //方式二导出word
	  function doExportWord2(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportUserWord2";
	  }
	  //方式一导出word（包含动态行表格）
	  function doExportWord3(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord3";
	  }
	  //方式一导出word（包含两个动态行表格）
	  function doExportWordD4(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWordD4";
	  }
	  //方式一导出word（包含动态行表格、循环列表中的动态行表格）
	  function doExportWord4(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord4";
	  }
	  //方式一导出word（合并单元格（合并列）--货物明细\人工费）
	  function doExportWord5(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord5";
	  }
	  //方式一导出word（合并单元格（一个列表下的合并行）--商品订单明细））
	  function doExportWord6(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord6";
	  }
	  
	  //方式一导出word（合并单元格（循环列表下的合并行）--商品订单明细）
	  function doExportWord7(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord7";
	  }
	  //方式一导出word（合并单元格（循环列表下的合并行）--商品订单明细、另加一个动态行表格）
	  function doExportWord8(){
	    window.location.href="<%=basePath%>/auth/exportWord/exportDataWord8";
	  }
	</script>
</body>
</html>



