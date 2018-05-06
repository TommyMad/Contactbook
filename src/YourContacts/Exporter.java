
package YourContacts;

import java.io.DataOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import javax.swing.JTable;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


/**
 *
 * @author Tomasz Madej
 */
public class Exporter {
    private File file;
    private List<JTable> table;
    private List<String> my_files;
    
    public Exporter (File file, List<JTable> table,List<String> my_files)
            throws Exception {
        this.file =file;
        this.table=table;
        this.my_files=my_files;
        
            if (my_files.size()!=table.size()){
            throw  new Exception ("Błąd");
            }
    }
    
    public boolean export(){
        try{
            DataOutputStream out=new DataOutputStream(new FileOutputStream(file));
            WritableWorkbook w=Workbook.createWorkbook(out);
            
            for (int index=0;index<table.size();index++){
                JTable new_table=table.get(index);
                WritableSheet s = w.createSheet(my_files.get(index),0);
                
                // creating table look in excel
                    for (int i=0;i<new_table.getColumnCount();i++){
                        for (int j=0;j<new_table.getRowCount();j++){
                            Object object=new_table.getValueAt(j, i);
                            s.addCell(new Label(0, 0, String.valueOf("Lp.")));
                            s.addCell(new Label(1, 0, String.valueOf("Imię")));
                            s.addCell(new Label(2, 0, String.valueOf("Nazwisko")));
                            s.addCell(new Label(3, 0, String.valueOf("Nr telefonu")));
                            s.addCell(new Label(4, 0, String.valueOf("E-mail")));
                            
                            s.addCell(new Label(i, j+1, String.valueOf(object)));
                        }
                    } 
            }
        w.write();
        w.close();
        return true;
       
        }
        catch(IOException | WriteException e){
            return false;
        }
    }

}
