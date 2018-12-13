/**
 *
 */
package com.james.exiapp.gui.pdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.util.PDFTextStripper;

/**
 * 解析英国皇家邮政的PDF
 *
 * @author TianYong
 *
 *
 */
//Tracked 48No Signature
//Delivered By:
//Postage on Account GB
//Safeplace
//32-033 388 5000-06E AB6 735
//MX 5161 9344 9GB
//SALON 66
//65 CHURCH STREET
//BURBAGE
//ITEM ID 1057129289
//HINCKLEY
//LE10 2DA
//Re
//tu
//rn
// A
//dd
//re
//ss
//
//Sm
//ar
//te
//ye
// C
//o.
//Lt
//d
//Un
//it
// D
//La
//ke
//si
//de
//Te
//ch
//no
//lo
//gy
// P
//ar
//k
//Ab
//er
//ta
//we
//SA
//7
//9F
//F
//Safeplace - 110 TABS TH
//Customer Reference:
//32820 / ORDER ID 1357117075
public class RoyalMail {

    public List<String> GetNoByWNo(InputStream pdfIn) throws Exception {
        //String msg = "导入PDF英国皇家邮政 ";
        //System.out.println(msg + " 开始");

        List<String> result = new ArrayList<String>();
        PDDocument doc = null;
        PDDocumentCatalog docCatalog=null;
        PDFTextStripper stripper=null;
        List<PDPage> pages = null;
        //PDPage page=null;
        String text=null;
        // String[] appId;

        try {
            // 读取导入的PDF文件
            doc = PDDocument.load(pdfIn);// 读取图像
            docCatalog = doc.getDocumentCatalog();
            stripper = new PDFTextStripper();
            stripper.setSortByPosition(false);

            pages = docCatalog.getAllPages();// 获取全部图像
            for (int i = 0; i < pages.size(); i++) {
                // 按页分解
                //page = pages.get(i);
                stripper.setStartPage(i + 1);
                stripper.setEndPage(i + 1);
                text = stripper.getText(doc);

                // int inx = text.indexOf("R\nE\n");
                text= text.replaceAll("\r", "");
                System.out.println(text);

                int ir = text.indexOf("Re\n");
                boolean falg = true;
                String tmp = text.substring(0);
                while(falg){

                    if(ir!=-1 && ir<text.length()){

                        String t = tmp.substring(ir).replaceAll("\n","");
                        if(t.startsWith("Return")){
                            falg = false;

                        }else{
                            //tmp = tmp.substring(ir+1);
                            ir+=1;
                        }
                    }else{
                        falg = false;
                    }
                }

                //text = text.substring(0,text.indexOf("R\nET\nU\nR\nN\n"));
                // if(inx==-1){
                //inx =   text.indexOf("R\nET\n");
                // }

//                                if(inx!=-1){
//                                    String tmpstr = text.substring(inx).replaceAll("\n", "").replaceAll(" ", "");
//                                    if(tmpstr.startsWith("RETURNTO")){
//                                         text = text.substring(0,inx);
//                                    }
//
//                                }

                result.add(text.substring(0,ir));
                // appId = text.split("Sender's reference:");
                // 如果数组不存在，则抛出pdf格式不正常，缺少Sender's reference:\n关键字
                // if (appId.length <= 0) {
                // throw new StorageException(
                // CodeEnum.WITHOUT_ROYALMAIL_PDF_REFERENCE, msg
                // + "pdf格式不正常，缺少Sender's reference:\n关键字");
                // }
                // result.put("" + appId[1].trim(),
                // getBytesFromPage(page.convertToImage()));
            }

            return result;
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (doc != null) {
                try {
                    doc.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * @param args
     */
    public static void main(String[] args) {
        RoyalMail royalMail = new RoyalMail();

        String fileName = "~/Downloads/GROUPON 0915 LABEL 1.pdf";
        File file = new File(fileName);
        InputStream pdfIn = null;
        try {
            if (file.exists()) {
                pdfIn = new FileInputStream(file);
                List<String> result = royalMail.GetNoByWNo(pdfIn);

                for (Object object : result) {
                    System.out.println(object);
                    break;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (pdfIn != null) {
                try {
                    pdfIn.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
