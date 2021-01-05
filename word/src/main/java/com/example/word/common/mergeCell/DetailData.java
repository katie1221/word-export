package com.example.word.common.mergeCell;

import java.util.ArrayList;
import java.util.List;

import com.deepoove.poi.data.RowRenderData;

public class DetailData {
	
    // 货品数据
    private List<RowRenderData> goods;
    
    // 人工费数据
    private List<RowRenderData> labors;

    public List<RowRenderData> getGoods() {
        return goods;
    }

    public void setGoods(List<RowRenderData> goods) {
        this.goods = goods;
    }

    public List<RowRenderData> getLabors() {
        return labors;
    }

    public void setLabors(List<RowRenderData> labors) {
        this.labors = labors;
    }

//	public List<RowRenderData> getGoods() {
//		List<RowRenderData> goods= new ArrayList<RowRenderData>();
//		for(int i=1;i<4;i++){
//			RowRenderData labor = RowRenderData.build("4", "墙纸", "书房+卧室", "1500", "/", "400", "1600");
//			goods.add(labor);
//		}
//		return goods;
//	}
//	public List<RowRenderData> getLabors() {
//		List<RowRenderData> labors= new ArrayList<RowRenderData>();
//		for(int i=1;i<4;i++){
//			RowRenderData labor = RowRenderData.build("油漆工", "2", "200", "400");
//			labors.add(labor);
//		}
//		return labors;
//	}
	
}
