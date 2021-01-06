package com.example.word.common.mergeCell3;


import java.util.List;
import java.util.Map;

import com.deepoove.poi.el.Name;

public class PaymentData3 {
	
    @Name("detail_table")
    private DetailData3 detailTable;
    @Name("typeLists")
    private List<Map<String,Object>> typeLists;//所属大类统计列表
    private String total;
    private String order_money;//订单总金额
    private String money_total;//订单总金额的中文大写
    
    public void setDetailTable(DetailData3 detailTable) {
        this.detailTable = detailTable;
    }

    public DetailData3 getDetailTable() {
        return this.detailTable;
    }

	public String getTotal() {
		return total;
	}

	public void setTotal(String total) {
		this.total = total;
	}

	public List<Map<String, Object>> getTypeLists() {
		return typeLists;
	}

	public void setTypeLists(List<Map<String, Object>> typeLists) {
		this.typeLists = typeLists;
	}

	public String getOrder_money() {
		return order_money;
	}

	public void setOrder_money(String order_money) {
		this.order_money = order_money;
	}

	public String getMoney_total() {
		return money_total;
	}

	public void setMoney_total(String money_total) {
		this.money_total = money_total;
	}
	
}
