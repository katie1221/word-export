package com.example.word.common.mergeCell2;

import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.util.TableTools;

/**
 * 表格动态行插入、渲染、合并单元格处理
 * @author Administrator
 *
 */
public class DetailTablePolicy2 extends DynamicTableRenderPolicy {

	// 填充数据所在行数
	int listsStartRow = 1;

	@Override
	public void render(XWPFTable table, Object data) {
		if (null == data) return;
		DetailData2 detailData = (DetailData2) data;

		// 商品订单详情列表数据 循环渲染
		List<RowRenderData> lists = detailData.getPlists();
		// 二级分类分组统计商品个数数据
		List<Map<String,Object>> tlists = detailData.getTlists();
		
		if (null != lists) {
			table.removeRow(listsStartRow);
			// 循环插入行
			for (int i = lists.size()-1; i >=0; i--) {
				XWPFTableRow insertNewTableRow = table.insertNewTableRow(listsStartRow);
				insertNewTableRow.setHeight(620);//设置行高
				// 循环插入列，共7列
				for (int j = 0; j < 7; j++)
					insertNewTableRow.createCell();
				// 渲染单行商品订单详情数据
				MiniTableRenderPolicy.Helper.renderRow(table, listsStartRow, lists.get(i));
			}
			//处理合并
			for (int i=0;i<lists.size();i++) {
				Object v =lists.get(i).getCells().get(1).getCellText();
				String type_name=String.valueOf(v);
				for(int j=0;j<tlists.size();j++){
					String typeName = String.valueOf(tlists.get(j).get("typeName"));
					Integer listSize = Integer.parseInt(String.valueOf(tlists.get(j).get("listSize")));
					if(type_name.equals(typeName)){
				        // 合并第1列的第i+1行到第i+listSize行的单元格
				        TableTools.mergeCellsVertically(table, 1, i+1, i+listSize);
//				        //处理垂直居中(默认行高时)
//				        for (int y = 0; y < 7; y++){
//					        XWPFTableCell cell = table.getRow(i+1).getCell(y);
//					        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); //垂直居中
//					    }
				        tlists.remove(j);
				        break;
					}
				}
				 //处理垂直居中(自定义行高时，设置垂直居中)
		        for (int y = 0; y < 7; y++){
			        XWPFTableCell cell = table.getRow(i+1).getCell(y);
			        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); //垂直居中
			    }
				System.out.println(v);
			}
		}
	}
}