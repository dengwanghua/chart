import java.awt.BasicStroke;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartFrame;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.axis.DateAxis;
import org.jfree.chart.axis.DateTickUnit;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.time.Millisecond;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.ui.TextAnchor;


public class MainFrame extends JFrame implements ActionListener{
	private JPanel contentPanel;
	private JPanel menuPanel;
	private JPanel functionPanel;
	public JButton paintButton;
	public JButton excelButton;
	public JLabel filepathLabel;
	public JTextField textField;
	public JLabel excelpathLabel;
	public JTextField exceltextField;
	public JButton saveButton;
	public JButton piantButton;
	public JButton excelchooseButton;
	public JLabel numberLabel;
	public JLabel daoLabel;
	public JTextField number1textField;
	public JTextField number2textField;
	public JButton chooseButton;
	public JButton startpaintButton;
	public JFreeChart chart[];
	public JPanel scrollJPanel;
	public JScrollPane scrollPane;
	public CategoryPlot plot;
	public Thread thread;
	public CategoryDataset[] dataset;
	public ChartPanel[] mChartFrame;
	public File file;
	public List<File> list;
	public int startnumber;
	public int stopnumber;
	public int chartwidth;
	int maxnumber;
	public MainFrame(){
		this.setTitle("绘制dat文件数据折线图");
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);   
		this.setSize(1000, 800);                          
		Dimension displaySize = Toolkit.getDefaultToolkit().getScreenSize();   
		Dimension frameSize = this.getSize();             
		if (frameSize.width > displaySize.width)  
			frameSize.width = displaySize.width;           
		if (frameSize.height > displaySize.height)  
			frameSize.height = displaySize.height;          
		this.setLocation((displaySize.width - frameSize.width) / 2,  
				(displaySize.height - frameSize.height) / 2);   
		//添加面板
		this.contentPanel=new JPanel();
		this.contentPanel.setBounds(0, 0, frameSize.width, frameSize.height);
		this.contentPanel.setBackground(new Color(255,255,255));
		this.contentPanel.setLayout(null);//设置布局NULL
		this.add(this.contentPanel);
		//功能菜单
		this.menuPanel=new JPanel();
		this.menuPanel.setBounds(10, 100, 140, 600);
		this.menuPanel.setBackground(new Color(255,255,255));
		this.menuPanel.setLayout(null);//设置布局NULL
		this.menuPanel.setBorder(BorderFactory.createLineBorder(Color.gray, 2, true));
		this.contentPanel.add(this.menuPanel);
		//绘制图像
		this.paintButton=new JButton(); 
		this.paintButton.setBounds(20, 150, 100, 30); 
		this.paintButton.addActionListener(this); 
		this.paintButton.setText("绘制图像");
        this.menuPanel.add(this.paintButton); 
		//文件数据截取
        this.excelButton=new JButton();
		this.excelButton.setBounds(20, 420, 100, 30); 
		this.excelButton.addActionListener(this); 
		this.excelButton.setText("截取文件");
        this.menuPanel.add(this.excelButton);
		 
		//功能内容区
		this.functionPanel=new JPanel(); 
		this.functionPanel.setLayout(null);//设置布局NULL
		this.functionPanel.setBounds(this.menuPanel.getX()+160, 50, 1000-this.menuPanel.getX()-180, 700);
		this.functionPanel.setBorder(BorderFactory.createLineBorder(Color.gray, 2, true));
		this.functionPanel.setBackground(new Color(255,255,255));
        this.contentPanel.add(this.functionPanel);
        //初始化控件
        this.filepathLabel=new JLabel("文件路径:");
    	this.filepathLabel.setBounds(20, 10, 70, 25);
    	
    	this.textField=new JTextField();
    	this.textField.enable(false);
		this.textField.setBounds(this.filepathLabel.getX()+this.filepathLabel.getWidth()+5, 10, 140, 25); 
		this.textField.setBorder(BorderFactory.createLineBorder(new Color(0, 0, 0)));
		
		this.chooseButton=new JButton("选择文件");  
		this.chooseButton.setBounds(this.textField.getX()+this.textField.getWidth()+5, 10, 100, 30); 
		this.chooseButton.addActionListener(this); 
		
		this.numberLabel=new JLabel();
		this.numberLabel.setBounds(this.chooseButton.getX()+this.chooseButton.getWidth()+5, 10, 50, 25); 
		this.numberLabel.setText("选择行:");
		
		this.number1textField=new JTextField();
		this.number1textField.setBounds(this.numberLabel.getX()+this.numberLabel.getWidth(), 10, 50, 25); 
		
		this.daoLabel=new JLabel("至",JLabel.CENTER);
		this.daoLabel.setBounds(this.number1textField.getX()+this.number1textField.getWidth(), 10, 25, 25);
		
		this.number2textField=new JTextField();
		this.number2textField.setBounds(this.daoLabel.getX()+this.daoLabel.getWidth(), 10, 50, 25); 
		
		this.startpaintButton=new JButton("开始绘图");  
		this.startpaintButton.setBounds(this.number2textField.getX()+this.number2textField.getWidth(), 10, 100, 30); 
		this.startpaintButton.addActionListener(this); 
		
		this.excelpathLabel=new JLabel("excel件路径:");
    	this.excelpathLabel.setBounds(20,10, 100, 25);
    	
    	this.exceltextField=new JTextField();
    	this.exceltextField.enable(false);
		this.exceltextField.setBounds(this.excelpathLabel.getX()+this.excelpathLabel.getWidth()+5, 10, 140, 25); 
		this.exceltextField.setBorder(BorderFactory.createLineBorder(new Color(0, 0, 0)));
		
		this.excelchooseButton=new JButton("选择excel文件");  
		this.excelchooseButton.setBounds(this.exceltextField.getX()+this.exceltextField.getWidth()+5, 10, 120, 30); 
		this.excelchooseButton.addActionListener(this); 
		
		this.saveButton=new JButton("截取");  
		this.saveButton.setBounds(this.excelchooseButton.getX()+this.excelchooseButton.getWidth(), 10, 100, 30); 
		this.saveButton.addActionListener(this); 
		
		addfunctionPanel();
		this.setVisible(true);
	}
	public void  addfunctionPanel(){
		this.functionPanel.removeAll();
		this.functionPanel.add(this.filepathLabel);
		this.functionPanel.add(this.textField);
		this.functionPanel.add(this.chooseButton);
		this.functionPanel.add(this.numberLabel); 
		this.functionPanel.add(this.number1textField);
		this.functionPanel.add(this.daoLabel);
		this.functionPanel.add(this.number2textField); 
		this.functionPanel.add(this.startpaintButton);
		this.functionPanel.repaint();
	}
	public void exceladdComponent(){
		this.functionPanel.removeAll();
		this.functionPanel.add(this.excelpathLabel);
		this.functionPanel.add(this.exceltextField);
		this.functionPanel.add(this.excelchooseButton);
		this.functionPanel.add(this.saveButton);
		this.functionPanel.repaint();
	}
	public  CategoryDataset[] GetDataset(List<File> files) throws Exception
	{
		chartwidth=this.functionPanel.getWidth()-10;
		DefaultCategoryDataset[] mDatasets=new DefaultCategoryDataset[191];
		Map<Integer, Map<Integer,Object>> maps = new HashMap<Integer, Map<Integer,Object>>(); 
		int z=1;
		for(File file:files){
			ReadExcelUtils excelReader = new ReadExcelUtils(file.getAbsolutePath()); 
			Map<Integer, Map<Integer,Object>> map = excelReader.readExcelContent();
			for (int i =1; i<= map.size(); i++) {
				maps.put(z++, map.get(i));
	        } 
		}
        for (int i =1; i<= maps.size(); i++) {
        	if(i>=10){
        		chartwidth=this.functionPanel.getWidth()-10;
        	}
        	 Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
        	 cellValue=maps.get(i);
        	 for (int j=0;j <cellValue.size(); j++) { 
        		 if(i==1){
        			 mDatasets[j]=new DefaultCategoryDataset();
        			 maxnumber=cellValue.size();
        		 }
        		 Double tmp=Double.parseDouble(cellValue.get(j).toString());
        		 mDatasets[j].addValue(tmp,String.valueOf(j), String.valueOf(i-1));
        	 }
        } 
		return mDatasets;
	}
	public void  paintChart(int i,int j){
		//定义长、宽、位置
		mChartFrame=new ChartPanel[191];
		chart=new JFreeChart[191];
		startnumber=i;
		stopnumber=j;
		int y=this.startpaintButton.getY()+this.startpaintButton.getHeight()+5;
		int width=this.functionPanel.getWidth()-10;
		int height=this.functionPanel.getHeight()-this.startpaintButton.getY()-this.startpaintButton.getHeight()-10;
		int myHeight=0;
		//创建容器
		this.scrollJPanel=new JPanel();
		this.scrollJPanel.setBackground(new Color(255,255,255));
		this.scrollJPanel.setLayout(null);
		//创建主题样式
        StandardChartTheme mChartTheme = new StandardChartTheme("CN");
        //设置标题字体
        mChartTheme.setExtraLargeFont(new Font("黑体", Font.BOLD, 15));
        //设置轴向字体
        mChartTheme.setLargeFont(new Font("宋体", Font.BOLD, 15));
        //设置图例字体
        mChartTheme.setRegularFont(new Font("宋体", Font.BOLD, 15));
        //应用主题样式
        ChartFactory.setChartTheme(mChartTheme);
		for(int k=startnumber-1;k<stopnumber;k++){
			chart[k] = ChartFactory.createLineChart("折线图", "时间", "数值", dataset[k], PlotOrientation.VERTICAL,true, true, false);
			CategoryPlot plot=chart[k].getCategoryPlot(); 
			plot.setBackgroundPaint(Color.LIGHT_GRAY);
			plot.setRangeGridlinePaint(Color.WHITE);
			plot.setRangeGridlinesVisible(false);
			plot.setDomainGridlinePaint(Color.WHITE);  
	        plot.setDomainGridlinesVisible(true); 
	        
		    LineAndShapeRenderer renderer=(LineAndShapeRenderer) plot.getRenderer();
		    renderer.setBaseItemLabelsVisible(true);
            renderer.setSeriesPaint(0, Color.black); 
		   
		    
			mChartFrame[k]= new ChartPanel(chart[k]);
			mChartFrame[k].setBounds(5, (k-startnumber+1)*(width/2+5), chartwidth,  width/2);
			mChartFrame[k].setVisible(true);
			myHeight+=5+width/2;
			this.scrollJPanel.add(mChartFrame[k]);
		}
		
		this.scrollJPanel.setPreferredSize(new Dimension(chartwidth+10,myHeight+10));
		this.scrollPane=new JScrollPane(this.scrollJPanel);
		this.scrollPane.setBounds(5,y,width, height);
		this.functionPanel.add(this.scrollPane);
		this.functionPanel.repaint();
		this.setVisible(true);
	}
	
	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		String source = e.getActionCommand();
		if(source.equals("绘制图像")){
			addfunctionPanel();
		}
		if(source.equals("截取文件")){
			exceladdComponent();
		}
		if(source.equals("选择excel文件")){
			JFileChooser jfc=new JFileChooser();  
	        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  
	        jfc.showDialog(new JLabel(), "选择文件");  
	        this.file=jfc.getSelectedFile(); 
	        if(file!=null){
	        	String fileName=file.getName();
		        if(file.isFile()&&(fileName.substring(fileName.lastIndexOf(".")+1).equals("xlsx")||fileName.substring(fileName.lastIndexOf(".")+1).equals("xls"))){
		        	this.exceltextField.setText(file.getAbsolutePath());
		        }
		        else{
		        	JOptionPane.showMessageDialog(null, "文件类型不正确，请选择.xlsx或.xls文件","提示",JOptionPane.ERROR_MESSAGE);
		        }
	        }
		}
		if(source.equals("截取")){
			String fileName=exceltextField.getText();
			if(fileName==null||fileName.length()<=0){
				JOptionPane.showMessageDialog(null,"请选择文件","提示",JOptionPane.ERROR_MESSAGE);
			}else{
				if(fileName.substring(fileName.lastIndexOf(".")+1).equals("dat")||fileName.substring(fileName.lastIndexOf(".")+1).equals("xls")||fileName.substring(fileName.lastIndexOf(".")+1).equals("xlsx")){
					 ReadExcelUtils excelReader = new ReadExcelUtils(file.getAbsolutePath());
				     try {
				    	 SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
						 String name=df.format(new Date());
						 String filestring=fileName.substring(0, fileName.lastIndexOf("/")+1)+name+".xlsx";
						 Workbook wb=new XSSFWorkbook();
						 Sheet sheet1=wb.createSheet("sheet0");;
						 Row row1;
						 FileOutputStream out=null;
						 Map<Integer, Map<Integer, Object>> map = excelReader.readExcelContent();
						 final Map<Integer, Map<Integer,Object>> newmap = new HashMap<Integer, Map<Integer,Object>>(); 
						 int z=1;
				         for (int i =1; i<= map.size(); i++) { 
				            	Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
				            	 cellValue=map.get(i);
				            		for (int j=0;j <cellValue.size(); j++) { 
					            		 String value=(String)cellValue.get(j);
					            			 if(value.indexOf("Df")!=-1){  
						            			 newmap.put(z++, map.get(i-1));
						            			 newmap.put(z++, map.get(i));
						            			 newmap.put(z++, map.get(i+1));		 
						            			 break;
					            	         } 
				            		 }
				            }  
				         if(newmap.size()>0){
				        	 for (int i =1; i<= newmap.size(); i++) { 
				        		 if(i%3==0){
				        			 row1=sheet1.createRow((short)(sheet1.getLastRowNum()+1)); //在现有行号后追加数据
					            	 Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
					            	 cellValue=newmap.get(i);
					            	 for (int j=0;j <cellValue.size(); j++) { 
					            		 String value=(String)cellValue.get(j);
							             row1.createCell(j).setCellValue(value);  
					            	 }
					            	 row1=sheet1.createRow((short)(sheet1.getLastRowNum()+1));
				        		 }else{
				        			 row1=sheet1.createRow((short)(sheet1.getLastRowNum()+1)); //在现有行号后追加数据
					            	 Map<Integer,Object> cellValue = new HashMap<Integer, Object>();
					            	 cellValue=map.get(i);
					            	 for (int j=0;j <cellValue.size(); j++) { 
					            		 String value=(String)cellValue.get(j);
							             row1.createCell(j).setCellValue(value);  
					            	 }
				        		 }
				        		
				            	
				            } 
				        	 out=new FileOutputStream(filestring);
			                 wb.write(out);
				        	 out.close();
				         }
				       
					} catch (Exception e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} 
					
					 
				}else{
					JOptionPane.showMessageDialog(null,"只支持dat、xlsx、xls格式文件，请重新选取","提示",JOptionPane.ERROR_MESSAGE);
				}
			}
		}
		if(source.equals("选择文件")){
			String path="";
			File[] files=null;
			list=null;
			list=new ArrayList<File>();
			ExcelFileFilter excelFilter = new ExcelFileFilter();
			JFileChooser jfc=new JFileChooser();  
	        jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
	        jfc.addChoosableFileFilter(excelFilter);
	        jfc.setFileFilter(excelFilter); 
	        jfc.setMultiSelectionEnabled(true);
	        jfc.showDialog(new JLabel(), "选择文件"); 
	        files=jfc.getSelectedFiles(); 
		    for(File f:files){
		    	list.add(f);
		    	path=f.getName()+" ";
		    }
		    Collections.sort(list);
		    textField.setText(path);
		}
		if(source.equals("开始绘图")){
		 maxnumber=0;
		 if(this.scrollPane!=null){
			 this.scrollJPanel.removeAll();
			 this.scrollPane.removeAll();
			 this.functionPanel.remove(this.scrollPane);
		 }
		 dataset=null;
		 mChartFrame=null;
		 chart=null;
		 if((this.number1textField.getText()==null||this.number1textField.getText().equals(""))||(this.number2textField.getText()== null||this.number2textField.getText().equals(""))){
				 JOptionPane.showMessageDialog(null, "行数不能为空,请重新选取","提示",JOptionPane.ERROR_MESSAGE);
		}else{
				 if(isNumeric(this.number1textField.getText())&&isNumeric(this.number2textField.getText())){
						int i=Integer.parseInt(this.number1textField.getText());
						int j=Integer.parseInt(this.number2textField.getText());
						if(i>=1&&i<=191&&j>=1&&j<=191){
							if(list.size()==0){
								JOptionPane.showMessageDialog(null,"请选择文件","提示",JOptionPane.ERROR_MESSAGE);
						    }else{
								    try {
								    	
								    	dataset=new CategoryDataset[191];
										dataset=this.GetDataset(list);
										if(j-i+1<=maxnumber){
											this.paintChart(i, j);
										}
										else{
											JOptionPane.showMessageDialog(null, "最多含有"+maxnumber+"列","提示",JOptionPane.ERROR_MESSAGE);
										}
										
									} catch (Exception e1) {
										// TODO Auto-generated catch block
										e1.printStackTrace();
									}
							}
						}
						else{
							JOptionPane.showMessageDialog(null, "选取的行数范围不对且前面行数要小于等于后面行数","提示",JOptionPane.ERROR_MESSAGE);
						}
						
					}
					else{
						JOptionPane.showMessageDialog(null, "选取的行数应为数字,请重新选取","提示",JOptionPane.ERROR_MESSAGE);
					}
		  }	
		}
	}
	public boolean isNumeric(String str){ 
		   Pattern pattern = Pattern.compile("[0-9]*"); 
		   Matcher isNum = pattern.matcher(str);
		   if( !isNum.matches() ){
		       return false; 
		   } 
		   return true; 
		}
	
}
