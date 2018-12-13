/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.james.exiapp.gui.model;

/**
 *
 * @author jameswee
 */
public class Vno {
    private String voucherCode;//VOUCHER CODE
    private String name;
    private String address;
    private String postcode;
    private String trackingCode;
    private String productName;
    private String icellV;
    private int rowNum;

    private String note;
    public Vno(int rowNum,String voucherCode, String name, String address, String postcode, String trackingCode, String note,String productName,String icellV) {
        this.rowNum = rowNum;
        this.voucherCode = voucherCode;
        this.name = name;
        this.address = address;
        this.postcode = postcode;
        this.trackingCode = trackingCode;
        this.note = note;
        this.productName = productName;
        this.icellV = icellV;
    }

    public Vno(){

    }
    public String getVoucherCode() {
        return voucherCode;
    }

    public void setVoucherCode(String voucherCode) {
        this.voucherCode = voucherCode;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getPostcode() {
        return postcode;
    }

    public void setPostcode(String postcode) {
        this.postcode = postcode;
    }

    public String getTrackingCode() {
        return trackingCode;
    }

    public void setTrackingCode(String trackingCode) {
        this.trackingCode = trackingCode;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public String getIcellV() {
        return icellV;
    }

    public void setIcellV(String icellV) {
        this.icellV = icellV;
    }

    public String vnoKeyByNameAndPostCode(){
        if(null==this.icellV){
            this.icellV = "";
        }
        //return
        return (this.postcode.toUpperCase()+this.icellV).replaceAll(" ", "").replaceAll("\r", "");
        //return this.getVoucherCode().trim();
    }
    @Override
    public String toString() {
        return "Vno{" + "voucherCode=" + voucherCode + ", name=" + name + ", address=" + address + ", postcode=" + postcode + ", trackingCode=" + trackingCode + ", productName=" + productName + ", rowNum=" + rowNum + ", note=" + note + '}';
    }


}
