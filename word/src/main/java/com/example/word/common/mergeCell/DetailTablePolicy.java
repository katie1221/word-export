package com.example.word.common.mergeCell;

import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.policy.DynamicTableRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.util.TableTools;

public class DetailTablePolicy extends DynamicTableRenderPolicy {

	  // 货品填充数据所在行数
	  int goodsStartRow = 2;
	  // 人工费填充数据所在行数
	  int laborsStartRow = 5;

	  @Override
	  public void render(XWPFTable table, Object data) {
	    if (null == data) return;
	    DetailData detailData = (DetailData) data;

	    // 人工费循环渲染
	    List<RowRenderData> labors = detailData.getLabors();
	    if (null != labors) {
	      table.removeRow(laborsStartRow);
	      // 循环插入行
	      for (int i = 0; i < labors.size(); i++) {
	        XWPFTableRow insertNewTableRow = table.insertNewTableRow(laborsStartRow);
	        for (int j = 0; j < 7; j++) insertNewTableRow.createCell();

	        // 合并单元格
	        TableTools.mergeCellsHorizonal(table, laborsStartRow, 0, 3);
	        // 渲染单行人工费数据
	        MiniTableRenderPolicy.Helper.renderRow(table, laborsStartRow, labors.get(i));
	      }
	    }

	    // 货品明细
	    List<RowRenderData> goods = detailData.getGoods();
	    if (null != goods) {
	      table.removeRow(goodsStartRow);
	      for (int i = 0; i < goods.size(); i++) {
	        XWPFTableRow insertNewTableRow = table.insertNewTableRow(goodsStartRow);
	        for (int j = 0; j < 7; j++) insertNewTableRow.createCell();
	        // 渲染单行货品明细数据
	        MiniTableRenderPolicy.Helper.renderRow(table, goodsStartRow, goods.get(i));
	      }
	    }
	  }
}