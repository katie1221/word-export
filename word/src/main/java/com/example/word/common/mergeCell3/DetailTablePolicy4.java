package com.example.word.common.mergeCell3;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.util.TableTools;

/**
 * 商品所属大类统计 表格动态行插入、渲染、合并单元格处理
 * @author Administrator
 *
 */
public class DetailTablePolicy4 extends DynamicTableRenderPolicy {

	// 填充数据所在行数
	int listsStartRow = 1;

	@Override
	public void render(XWPFTable table, Object data) {
		if (null == data) return;
		//文本居中
		TableStyle rowStyle = new TableStyle();
		rowStyle = new TableStyle();
		rowStyle.setAlign(STJc.CENTER);
		
		List<RowRenderData> lists = new ArrayList<RowRenderData>();
		//所属大类列表数据
		List<Map<String,Object>> typeLists = (List<Map<String, Object>>) data;
        for(Map<String,Object> map:typeLists){
        	String index =String.valueOf(map.get("index")); 
        	String sub_type =String.valueOf(map.get("sub_type")); 
        	String total_price =String.valueOf(map.get("total_price")); 
        	RowRenderData tlist = RowRenderData.build(index, sub_type, total_price,total_price);
        	tlist.setRowStyle(rowStyle);
			lists.add(tlist);
        }
		
		if (null != lists) {
			table.removeRow(listsStartRow);
			// 循环插入行
			for (int i = lists.size()-1; i >=0; i--) {
				XWPFTableRow insertNewTableRow = table.insertNewTableRow(listsStartRow);
				// 循环插入列，共4列
				for (int j = 0; j < 4; j++)
					insertNewTableRow.createCell();
				// 渲染单行所属大类数据
				MiniTableRenderPolicy.Helper.renderRow(table, listsStartRow, lists.get(i));
			}
		}
	}
}