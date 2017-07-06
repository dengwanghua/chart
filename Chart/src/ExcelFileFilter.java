import java.io.File;

import javax.swing.filechooser.FileFilter;


public class ExcelFileFilter extends FileFilter{

	@Override
	public boolean accept(File f) {
		String name = f.getName();    
        return  name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx")||name.toLowerCase().endsWith(".dat");
	}

	@Override
	public String getDescription() {
		return   "*.xls;*.xlsx;*.dat";
	}

}
