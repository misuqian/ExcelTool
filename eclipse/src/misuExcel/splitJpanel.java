/*
 Copyright (c) 2016 pengqian
 
 Permission is hereby granted, free of charge, to any person obtaining a copy
 of this software and associated documentation files (the "Software"), to deal
 in the Software without restriction, including without limitation the rights
 to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is
 furnished to do so, subject to the following conditions:
 
 The above copyright notice and this permission notice shall be included in all
 copies or substantial portions of the Software.
 
 THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 SOFTWARE.
 */


package misuExcel;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.File;
import java.util.ArrayList;

import javax.swing.Box;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JSpinner;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;

import org.apache.poi.ss.usermodel.Workbook;

public class splitJpanel implements ActionListener,ItemListener{
	private JPanel jp01 =null;
	private JTextField address01,address02,name01,saveAddress,name02;
	private JSpinner ignoreRow,ignoreCell,ignoreRow_tar,ignoreCell_tar;
	private JButton setButton01,setButton02,setSave,startButton;
	private Box selectBox01,selectBox02;
	private JComboBox sheet_targetBox,cell_targetBox,sheet_excelBox,cell_excelBox,cell_targetBox2;
	private excelRead read01,read02;
	private Workbook excel01,excel02;
	private JRadioButton r1,r2,r3;
	public static boolean readyEx01=false;
	public static boolean readyEx02=false;
	private int type = 1;
	private int typeCankao = 1;
	public static int ignore_Row = 0;
	public static int ignore_Rowtar = 0;
	private boolean started =false;
	
	public splitJpanel(){
		initJpanel01();	
	}
	
	private void initJpanel01(){
		jp01 = new JPanel();
		JLabel title = new JLabel("①请选择分割目标 :");
//		JLabel abLabel = new JLabel("读取路径 :");
//		JLabel naLabel = new JLabel("读取文件名 :");
		address01 = new JTextField(100);	address01.setEditable(false);
		name01 = new JTextField(50);	name01.setEditable(false); 
		setButton01 = new JButton("选择Excel文件"); setButton01.addActionListener(this);
		
		Box box01 = Box.createHorizontalBox();
		box01.add(new JLabel("	读取路径"));  box01.add(Box.createHorizontalStrut(15)); box01.add(address01);	 box01.add(Box.createHorizontalStrut(50));
		Box box02 = Box.createHorizontalBox();
		box02.add(new JLabel("	读取名字"));  box02.add(Box.createHorizontalStrut(15)); box02.add(name01);	 box02.add(Box.createHorizontalStrut(50));
		
		JLabel title2 = new JLabel("②请选择参考文件 :");
//		JLabel abLabel2 = new JLabel("读取路径 :");
//		JLabel naLabel2 = new JLabel("读取文件名 :");
		address02 = new JTextField(100);	address02.setEditable(false);
		name02 = new JTextField(50);	name02.setEditable(false); 
		JLabel abLabel3 = new JLabel("③请选择分割文件保存路径 :");
		saveAddress = new JTextField(100);	saveAddress.setEditable(false);
		setButton02 = new JButton("选择Excel文件"); 	setButton02.addActionListener(this);
		ButtonGroup group =new ButtonGroup();
		r1= new JRadioButton("分类到不同xls文件"); 
		r1.setSelected(true); r1.addItemListener(this); 
		r2= new JRadioButton("分类到不同表格");
		group.add(r1); group.add(r2);	r2.addItemListener(this);
		
		
		Box box03 = Box.createHorizontalBox();
		box03.add(new JLabel("	读取路径")); box03.add(Box.createHorizontalStrut(15));box03.add(address02);box03.add(Box.createHorizontalStrut(50));
		Box box04 = Box.createHorizontalBox();
		box04.add(new JLabel("	读取名字")); box04.add(Box.createHorizontalStrut(15));box04.add(name02);box04.add(Box.createHorizontalStrut(50));
		Box boxTaggle = Box.createHorizontalBox();
		boxTaggle.add(new JLabel("	分割类型")); boxTaggle.add(Box.createHorizontalStrut(10)); boxTaggle.add(r1); boxTaggle.add(r2);
		
		Box taggleBox02 = Box.createHorizontalBox();
		r3 = new JRadioButton("列参考");
		taggleBox02.add(new JLabel("	分割方式")); taggleBox02.add(Box.createHorizontalStrut(10)); taggleBox02.add(r3); r3.setSelected(true); r3.setEnabled(false);
		
		setSave = new JButton("选择路径"); setSave.addActionListener(this);
		
		Box savebox = Box.createHorizontalBox();
		savebox.add(Box.createHorizontalStrut(15));
		savebox.add(saveAddress); savebox.add(Box.createHorizontalStrut(15));savebox.add(setSave);savebox.add(Box.createHorizontalStrut(15));
		
		Box ignoreBox = Box.createHorizontalBox();
		ignoreRow = new JSpinner(new SpinnerNumberModel(0,0,100,1));
		ignoreCell = new JSpinner(new SpinnerNumberModel(0,0,100,1));
		ignoreBox.add(new JLabel(" 忽略行数"));  ignoreBox.add(Box.createHorizontalStrut(15)); ignoreBox.add(ignoreRow);
		ignoreBox.add(Box.createHorizontalStrut(30));
		ignoreBox.add(new JLabel("  忽略列数")); ignoreBox.add(Box.createHorizontalStrut(15));	ignoreBox.add(ignoreCell); ignoreCell.setEnabled(false);
		ignoreBox.add(Box.createHorizontalStrut(300));
		
		Box ignoreBox_tar = Box.createHorizontalBox();
		ignoreRow_tar = new JSpinner(new SpinnerNumberModel(0,0,100,1));
		ignoreCell_tar = new JSpinner(new SpinnerNumberModel(0,0,100,1));
		ignoreBox_tar.add(new JLabel(" 忽略行数")); ignoreBox_tar.add(Box.createHorizontalStrut(15));ignoreBox_tar.add(ignoreRow_tar);
		ignoreBox_tar.add(Box.createHorizontalStrut(30));
		ignoreBox_tar.add(new JLabel("  忽略列数")); ignoreBox_tar.add(Box.createHorizontalStrut(15));ignoreBox_tar.add(ignoreCell_tar); 
		ignoreBox_tar.add(Box.createHorizontalStrut(300));
		ignoreCell_tar.setEnabled(false);
		
		selectBox01 = Box.createHorizontalBox();
		selectBox01.setVisible(false);
		sheet_excelBox =new JComboBox(new DefaultComboBoxModel());		sheet_excelBox.addActionListener(this);
		cell_excelBox = new JComboBox(new DefaultComboBoxModel());		
		selectBox01.add(new JLabel(" 参考工作表:"));	 selectBox01.add(Box.createHorizontalStrut(15)); selectBox01.add(sheet_excelBox);
		selectBox01.add(Box.createHorizontalStrut(30)); 
		selectBox01.add(new JLabel(" 参考列"));	 selectBox01.add(Box.createHorizontalStrut(15)); 
		selectBox01.add(cell_excelBox);	 selectBox01.add(Box.createHorizontalStrut(300));
		
		selectBox02 = Box.createHorizontalBox();
		selectBox02.setVisible(false);  //选择之后弹出
//		JLabel boxjl01 = new JLabel("参考工作表:");
		sheet_targetBox =new JComboBox(new DefaultComboBoxModel());		sheet_targetBox.addActionListener(this);
//		JLabel boxjl02 = new JLabel(" 参考列:");
		cell_targetBox = new JComboBox(new DefaultComboBoxModel());	
		selectBox02.add(new JLabel(" 参考工作表:")); selectBox02.add(Box.createHorizontalStrut(15)); selectBox02.add(sheet_targetBox); 
		selectBox02.add(Box.createHorizontalStrut(30));
		selectBox02.add(new JLabel(" 参考列")); selectBox02.add(Box.createHorizontalStrut(15)); selectBox02.add(cell_targetBox);
		selectBox02.add(Box.createHorizontalStrut(300));
		
		startButton = new JButton("开始分割");
		startButton.addActionListener(this);
		
		
		jp01.setLayout(new GridLayout(17,1,0,5));
		jp01.add(title);  jp01.add(taggleBox02); jp01.add(ignoreBox);
		jp01.add(box02); jp01.add(box01); 	jp01.add(setButton01); 
		jp01.add(selectBox01);
//		jp01.add(new JLabel("	-----------------------------------------------------------------------	"));
		jp01.add(title2); jp01.add(boxTaggle);	jp01.add(ignoreBox_tar);
		jp01.add(box04); jp01.add(box03);	jp01.add(setButton02);
		jp01.add(selectBox02);
//		jp01.add(new JLabel("	-----------------------------------------------------------------------	"));
		jp01.add(abLabel3);	 jp01.add(savebox);
//		jp01.add(saveAddress); 	jp01.add(setSave);
		jp01.add(startButton);
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		if(e.getSource()==setButton01){
			final File file = showChooser(JFileChooser.FILES_ONLY);
			if((file!=null)&&(file.isFile())){
				if(examExcel(file.getName())){
						address01.setText(file.getAbsolutePath());
						name01.setText(file.getName());
						new Thread(){
							@Override
							public void run() {
								super.run();
								readExcel(file.getPath(),1);
								showBox(1);
							}
						}.start();
						//选择成功～～～～～～～～～～～～～～～～～～～～～
						 started =false;
				}else{
					JOptionPane.showMessageDialog(null,"请选择.xls或者.xlsx格式文件");
				}
			}
		}//end 01 if
		else if(e.getSource()==setButton02){
			final File file = showChooser(JFileChooser.FILES_ONLY);
			if((file!=null)&&(file.isFile())){
				if(examExcel(file.getName())){	
					//选择成功～～～～～～～～～～～～～～～～～～～～～
					name02.setText(file.getName());
					address02.setText(file.getAbsolutePath());
					new Thread(){
						@Override
						public void run() {
							super.run();
							readExcel(file.getPath(), 2);
							showBox(2);
						}
					}.start();
					 started =false;
				}else{
					JOptionPane.showMessageDialog(null,"请选择.xls或者.xlsx格式文件");
				}
			}
		}//end 02 if
		else if (e.getSource()==setSave){
			File file = showChooser(JFileChooser.DIRECTORIES_ONLY);
			if((file!=null)&&(file.isDirectory())){
				saveAddress.setText(file.getAbsolutePath());
				 started =false;
			}
		}
		
		else if (e.getSource()==sheet_targetBox){
//			setCellBox()
			if(excel02!=null&&read02!=null){
				setCellBox(read02.getSCellNum(sheet_targetBox.getSelectedIndex()),2);
				 started =false;
			}
		}else if(e.getSource()==sheet_excelBox){
			if(excel01!=null&&read01!=null){
				setCellBox(read01.getSCellNum(sheet_excelBox.getSelectedIndex()),1);
				 started =false;
			}
		}
		else if(e.getSource()==startButton){
		if(!started){	
			started=true;
//			if(ignoreRow.getText().trim().matches("\\d+")
//						&&ignoreRow_tar.getText().trim().matches("\\d+")){
			if(excel01!=null&&excel02!=null){
				new Thread(){
					@Override
					public void run() {
						super.run();	
						Log.init(saveAddress.getText().trim());
						Log.info("开始分割");
//						ignore_Row = Integer.parseInt(ignoreRow.getText().trim());
//						ignore_Rowtar = Integer.parseInt(ignoreRow_tar.getText().trim());
						ignore_Row = (int) ignoreRow.getValue();
						ignore_Rowtar =(int) ignoreRow_tar.getValue();
						excelSplit split = new excelSplit(excel01, excel02,sheet_excelBox.getSelectedIndex(), cell_excelBox.getSelectedIndex(), 
								sheet_targetBox.getSelectedIndex(), cell_targetBox.getSelectedIndex(),
								type,saveAddress.getText().trim(),name01.getText().trim());
						split.splitExcel();
					}
				}.start();
			}else {
				 JOptionPane.showMessageDialog(null,"请选择拼接母体或参考文件");
			}
//		 }else{
//			 JOptionPane.showMessageDialog(null,"忽略行列请输入有效数字");
//		 }
		}
		}
	}
	
	@Override
	public void itemStateChanged(ItemEvent e) {
		if(r1.isSelected()){
			type =1;
		}else if(r2.isSelected()){
			type =2 ;
		}		
	}
	
	private File showChooser(int mode){
		JFileChooser jfc=new JFileChooser();
		jfc.setFileSelectionMode(mode);
		jfc.showOpenDialog(null);
		return jfc.getSelectedFile();
	}
	
	private boolean examExcel(String name){
		String[] nameStr = name.split("\\.");
		return (nameStr[nameStr.length-1].equals("xls")||nameStr[nameStr.length-1].equals("xlsx"));
	}
	
	
	private void readExcel(String path,int type){
		if(type==1){
			read01 = new excelRead(path,type);
			excel01= read01.getWorkbook();
		}
		else if(type==2){
			read02 = new excelRead(path,type);
			excel02 = read02.getWorkbook();
		}
		else 
			System.out.println("代码错误");
	}
	
	public void showBox(int type){
	 switch (type) {
		case 1:
			selectBox01.setVisible(true);
			if(excel01!=null&&read01!=null){
				 setSheetBox(read01.getSheetNames(),1);
			}else{
//				System.out.println("sheet读取错误");
			}
			break;
		case 2:
			selectBox02.setVisible(true);
			if(excel02!=null&&read02!=null){
				 setSheetBox(read02.getSheetNames(),2);
			}else{
//				System.out.println("sheet读取错误");
			}
			break;
		default:
			break;
		}	
	}
	
	private void setSheetBox(ArrayList<String> list,int type){
		switch (type) {
		case 1:
			sheet_excelBox.removeAllItems();
			for(int i=0;i<list.size();i++){
				sheet_excelBox.addItem(list.get(i)+" 表");
			}
			break;
		case 2:
			sheet_targetBox.removeAllItems();
			for(int i=0;i<list.size();i++){
				sheet_targetBox.addItem(list.get(i)+" 表");
			}
			break;
		default:
			break;
		}
	}
	
	private void setCellBox(int count,int type){
		switch (type) {
		case 1:
			cell_excelBox.removeAllItems();
			for(int i =0 ; i < count ; i++){
				cell_excelBox.addItem("第"+(i+1)+"列");
			}
			break;
		case 2:
			cell_targetBox.removeAllItems();
			for(int i =0 ; i < count ; i++){
				cell_targetBox.addItem("第"+(i+1)+"列");
			}
			break;
		default:
			break;
		}
	}
	
	public JPanel getPanel(){
		return jp01;
	}

}
