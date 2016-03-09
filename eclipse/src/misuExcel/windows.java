package misuExcel;

import java.awt.BorderLayout;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;


public class windows{
//	private static final int String = 0;
	private JFrame jFrame;
	private JTabbedPane jtab;
	private JPanel jp01,jp02,jp03;
	
	
	public windows(){
		initWindow();
	}
	
	private void initWindow(){
		jFrame = new JFrame("ExcelTool");
		splitJpanel splitJpanel = new splitJpanel();
		jp02 = splitJpanel.getPanel();
//		spliceJpanel spliceJpanel = new spliceJpanel();
		addJpanel addJpanel  = new addJpanel();
		jp01 = addJpanel.getPanel();
		
		initJp03();
		
		jtab = new JTabbedPane();
		jtab.setPreferredSize(new Dimension(800,600));
	if(jp01!=null)
		jtab.add(jp01,"智能拼接");
	if(jp02!=null)	
		jtab.add(jp02,"智能分割");
	if(jp03!=null)
		jtab.add(jp03,"帮助");
	
		jFrame.setBounds(250,100,800,600);
		jFrame.setResizable(false);
		jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		jFrame.getContentPane().add(jtab,BorderLayout.CENTER);
		jFrame.setVisible(true);
	}
	
	private void initJp03(){
		jp03 = new JPanel();
		JLabel jl = new JLabel("ExcelTool是快速拼接，分割纯文本Excel的软件",JLabel.CENTER);
		JLabel jl02 = new JLabel("使用软件操作Excel时候请务必保证excel没有合并项，未使用函数，为纯文本格式",JLabel.CENTER);
		JButton jb = new JButton("查看帮助PDF");
		jb.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
//				System.out.println("./pdf/ExcelTool.pdf");
//				run.OStype==1?"./src/pdf/ExcelTool.pdf":".\\src\\pdf\\ExcelTool.pdf"
//				windows.class.getClassLoader().getResource(run.OStype==1?"pdfread/ExcelTool.pdf":"pdfread\\ExcelTool.pdf").getFile()
					try {	
							java.net.URI uri = new java.net.URI("https://github.com/misuqian");
							Desktop.getDesktop().browse(uri);
//						Desktop.getDesktop().open(
//								new File(run.OStype==1?"./ExcelTool.pdf":".\\ExcelTool.pdf"));
					}catch (URISyntaxException e1) {
//						e1.printStackTrace();
					}catch (IOException e1) {
//						e1.printStackTrace();
					}
			}
		});
		jp03.add(jl,BorderLayout.NORTH);
		jp03.add(jl02,BorderLayout.NORTH);
		jp03.add(jb,BorderLayout.CENTER);
	}
	
}
