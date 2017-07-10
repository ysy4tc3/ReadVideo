import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class ReadVideo {
    public static void main(String[] args){
        try {
            ArrayList<File> fl = ReadFile.getFile(args[0]);
            ExcelOperation.createExcel(new FileOutputStream(args[1]+"log.xls"),fl);
        }catch(FileNotFoundException e){
            e.printStackTrace();
        }catch(Exception e) {
            e.printStackTrace();
        }
    }
}

class ReadFile{
    /**
     * 读取某个文件夹下的所有文件
     */
    public static ArrayList<File> getFile(String stringPath) throws FileNotFoundException{
        File f = new File(stringPath);
        if(!f.exists()){
            throw new FileNotFoundException(stringPath +": File not exist");
        }

        File[] fa = f.listFiles();
        ArrayList<File> fl = new ArrayList<File>();
        for(int i = 0; i < fa.length; i++){
            if(!fa[i].isDirectory()){
                fl.add(fa[i]);
            }
        }
        return fl;
    }
    /**
     * 读取某个文件夹下的所有文件夹
     */
    public static ArrayList<File> getDirectory(String stringPath) throws FileNotFoundException{
        File f = new File(stringPath);
        if(!f.exists()){
            throw new FileNotFoundException(stringPath +": File not exist");
        }

        File[] fa = f.listFiles();
        ArrayList<File> fl = new ArrayList<File>();
        for(int i = 0; i < fa.length; i++){
            if(fa[i].isDirectory()){
                fl.add(fa[i]);
            }
        }
        return fl;
    }
}