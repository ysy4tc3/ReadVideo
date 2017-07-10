import java.io.IOException;
import java.io.OutputStream;

import jxl.Workbook;
import jxl.write.*;

import it.sauronsoftware.jave.Encoder;
import it.sauronsoftware.jave.MultimediaInfo;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

public class ExcelOperation {
    public static String[] alphabet = {
            "A", "B", "C", "D", "E",
            "F", "G", "H", "I", "J",
            "K", "L", "M", "N", "O",
            "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y",
            "Z"
    };
    public static void createExcel(OutputStream os, ArrayList<File> fl) throws WriteException, IOException {
        System.out.print("Progress:");
        //创建工作薄
        WritableWorkbook workbook = Workbook.createWorkbook(os);
        //创建新的一页
        WritableSheet sheet = workbook.createSheet("First Sheet", 0);
        //创建要显示的内容,创建一个单元格，第一个参数为列坐标，第二个参数为行坐标，第三个参数为内容
        Label dateLabel = new Label(0, 0, "日期");
        sheet.addCell(dateLabel);
        Label fileName = new Label(1, 0, "文件名");
        sheet.addCell(fileName);
        Label duration = new Label(2, 0, "时长");
        sheet.addCell(duration);

        File source;
        Encoder encoder = new Encoder();
        String date = "";
        String rgex = "WIN_([0-9]{8}_[0-9]{2}_[0-9]{2}_[0-9]{2})";
        String rgex2 = "WIN_([0-9]{8})";
        int size = fl.size();
        System.out.print("0/"+size);
        for (int i = 0; i < size; i++) {
            try {
                source = fl.get(i);
                MultimediaInfo m = encoder.getInfo(source);
                long ls = m.getDuration();
                String name = source.getName();

                if (!DealStrSub.getSubUtilSimple(date, "([0-9]{8})").equals(DealStrSub.getSubUtilSimple(name, rgex2))) {
                    date = DealStrSub.getSubUtilSimple(name, rgex);
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HH_mm_ss");
                    Date realDate = sdf.parse(date);
                    WritableCellFormat wcf = new WritableCellFormat(DateFormats.FORMAT1);
                    DateTime aDate = new DateTime(0, i + 1, realDate, wcf);
                    sheet.addCell(aDate);
                }
                Label aName = new Label(1, i + 1, name);
                sheet.addCell(aName);
                Label aDuration = new Label(2, i + 1, "" + ls / 3600000 + ":" + ls % 3600000 / 60000 + ":" + (ls % 60000) / 1000 );
                sheet.addCell(aDuration);

                if(i == 1){
                    Formula form = new Formula(3, i+1,"TEXT(C2+C3,\"[h]:m:s\"",new WritableCellFormat(new NumberFormat("#.00")));
                    sheet.addCell(form);
                }else{
                    Formula form = new Formula(3,i+1,"TEXT(C"+(i+2)+"+D"+(i+1)+",\"[h]:m:s\"");
                    sheet.addCell(form);
                }

                Formula x = new Formula(3, 1, "C2");
                sheet.addCell(x);

            } catch (Exception e) {
                e.printStackTrace();
            }
            Label label = new Label(0,size+1,"总时长");
            sheet.addCell(label);
            Formula form = new Formula(2, size+1,"D"+(size+1));
            sheet.addCell(form);
            System.out.print("\n");
            System.out.print(""+(i+1)+"/"+size);
        }

        //把创建的内容写入到输出流中，并关闭输出流
        workbook.write();
        workbook.close();
        os.close();
    }
}
    