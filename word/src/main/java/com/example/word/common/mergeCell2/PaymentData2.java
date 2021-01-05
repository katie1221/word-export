package com.example.word.common.mergeCell2;

import com.deepoove.poi.el.Name;

public class PaymentData2 {
	
    private String NO;
    private String ID;
    private String taitou;
    private String consignee;
    @Name("detail_table")
    private DetailData2 detailTable;
    private String subtotal;
    private String tax;
    private String transform;
    private String other;
    private String unpay;
    private String total;


    public void setNO(String NO) {
        this.NO = NO;
    }

    public String getNO() {
        return this.NO;
    }

    public void setID(String ID) {
        this.ID = ID;
    }

    public String getID() {
        return this.ID;
    }

    public void setTaitou(String taitou) {
        this.taitou = taitou;
    }

    public String getTaitou() {
        return this.taitou;
    }

    public void setConsignee(String consignee) {
        this.consignee = consignee;
    }

    public String getConsignee() {
        return this.consignee;
    }

    public void setDetailTable(DetailData2 detailTable) {
        this.detailTable = detailTable;
    }

    public DetailData2 getDetailTable() {
        return this.detailTable;
    }

    public void setSubtotal(String subtotal) {
        this.subtotal = subtotal;
    }

    public String getSubtotal() {
        return this.subtotal;
    }

    public void setTax(String tax) {
        this.tax = tax;
    }

    public String getTax() {
        return this.tax;
    }

    public void setTransform(String transform) {
        this.transform = transform;
    }

    public String getTransform() {
        return this.transform;
    }

    public void setOther(String other) {
        this.other = other;
    }

    public String getOther() {
        return this.other;
    }

    public void setUnpay(String unpay) {
        this.unpay = unpay;
    }

    public String getUnpay() {
        return this.unpay;
    }

    public String getTotal() {
        return total;
    }

    public void setTotal(String total) {
        this.total = total;
    }
}
