/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.james.exiapp.gui.excel;

import com.james.exiapp.gui.model.Vno;
import com.james.exiapp.gui.pdf.RoyalMail;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author jameswee
 */
public class ExcelUtil {

    /**
     * 读取EXCEL
     *
     * @param filePath
     * @return
     * @throws Exception
     */
    public List<Vno> readExcel(String filePath) throws Exception {
        InputStream input = new FileInputStream(filePath);
        List<Vno> list = new ArrayList<Vno>();

        try {
            Workbook wb = WorkbookFactory.create(input);
            Sheet sheet = wb.getSheetAt(0);
            java.util.Iterator<Row> rit = sheet.rowIterator();
            rit.next();
            Vno eor = null;
            Cell trcode = null;
            Row row = null;
            for (; rit.hasNext();) {
                row = rit.next();

                eor = new Vno();
                eor.setRowNum(row.getRowNum() + 1);

                trcode = row.getCell(0, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setVoucherCode(trcode.getStringCellValue());
                } else {
                    eor.setVoucherCode(null);
                }
                trcode = row.getCell(1, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setName(trcode.getStringCellValue());
                } else {
                    eor.setName(null);
                }
                trcode = row.getCell(2, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setAddress(trcode.getStringCellValue());
                } else {
                    eor.setAddress(null);
                }
                trcode = row.getCell(3, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setPostcode(trcode.getStringCellValue());
                } else {
                    eor.setPostcode(null);
                }
                trcode = row.getCell(4, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setTrackingCode(trcode.getStringCellValue());
                } else {
                    eor.setTrackingCode(null);
                }
                trcode = row.getCell(5, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setNote(trcode.getStringCellValue());
                } else {
                    eor.setNote(null);
                }

                trcode = row.getCell(6, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setProductName(trcode.getStringCellValue());
                } else {
                    eor.setProductName(null);
                }
                trcode = row.getCell(8, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    eor.setIcellV(trcode.getStringCellValue());
                } else {
                    eor.setIcellV(null);
                }

                list.add(eor);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();

        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 更新EXCEL
     *
     * @param filePath
     * @param vnoList
     * @throws Exception
     */
    public String writeExcel(String excelPath, String pdfPath, List<Vno> vnoList) throws Exception {
        File excelFile = new File(excelPath);
        Map<Integer, Vno> vnoMap = new HashMap<Integer, Vno>();
        Map<String, List<String>> pdfMap = new HashMap<String, List<String>>();

        //List<?> pdfList = null;
        try {
            InputStream pdfInput = new FileInputStream(pdfPath);
            RoyalMail royal = new RoyalMail();
            List<String> pdfList = royal.GetNoByWNo(pdfInput);
            List<String> tmp = null;
            String tmpKey = null;
            for (String pdf : pdfList) {
                tmp = Arrays.asList(pdf.split("\n"));
                //tmpKey=tmp.get(tmp.size()-3).trim().replaceAll(" ", "");
                int len = tmp.size() - 1;
                tmpKey = tmp.get(len);
                if (tmp.size() >= 4) {
                    String str = tmp.get(len - 2);

                    str = null != str ? str.replaceAll(" ", "") : "";
                    if (str.startsWith("ITEMID")) {
                        tmpKey += (str.replaceAll("ITEMID", ""));
                    }
                    //else {
                    //                      str = tmp.get(10);
                    //                       str = null!=str?str.replaceAll(" ",""):"";
                    //                      tmpKey+=(str.replaceAll("ITEMID", ""));
                    //                  }
                }
                tmpKey = (tmpKey.replaceAll(" ", "").toUpperCase()).replaceAll("\r", "");
//           System.out.println("tmpKey=>"+tmpKey);
                pdfMap.put(tmpKey, tmp);
//              if(tmpKey.contains(",")){
//                  String[] vs = tmpKey.split(",");
//                  for(String v : vs){
//                       pdfMap.put(v, tmp);
//                  }
//              }else{
//              //tmpkey = tmp.get(tmp.size()-3);
//               pdfMap.put(tmpKey, tmp);
//              }
            }

            String vnokey = null;
            for (Vno v : vnoList) {
                vnokey = v.vnoKeyByNameAndPostCode();
//               System.out.println("vnokey=>"+vnokey);
                tmp = pdfMap.get(vnokey);
                if (null != tmp) {
                    v.setTrackingCode(tmp.get(5).replaceAll(" ", ""));
                }

                vnoMap.put(Integer.valueOf(v.getRowNum()), v);
            }
//

            InputStream input = new FileInputStream(excelFile);
            Workbook wb = WorkbookFactory.create(input);
            Sheet sheet = wb.getSheetAt(0);
            java.util.Iterator<Row> rit = sheet.rowIterator();
            rit.next();
            Vno eor = null;
            Cell trcode = null;
            Row row = null;
            for (; rit.hasNext();) {
                row = rit.next();
                Integer index = row.getRowNum() + 1;
                eor = vnoMap.get(index);
                trcode = row.getCell(4, Row.CREATE_NULL_AS_BLANK);
                if (null != trcode && null != eor) {
                    trcode.setCellType(HSSFCell.CELL_TYPE_STRING);
                    trcode.setCellValue(eor.getTrackingCode());
                }
            }
            // finshFile = excelFile.getParent()+File.separator+ "update_finish_"+excelFile.getName();
            OutputStream fileOut = new FileOutputStream(excelFile);
            wb.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            return null;
        }
        return excelFile.getPath();
    }

}
