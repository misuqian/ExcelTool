package misuExcel;

import java.util.Properties;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;


public class run {
	public static int OStype = 1;
	
	public static void main(String[] args) {
//		try {
//			BeautyEyeLNFHelper.frameBorderStyle = BeautyEyeLNFHelper.FrameBorderStyle.generalNoTranslucencyShadow;
//			UIManager.put("RootPane.setupButtonVisible", false);
//			UIManager.put("TabbedPane.tabAreaInsets"
//				    , new javax.swing.plaf.InsetsUIResource(3,20,2,20));
//			org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper.launchBeautyEyeLNF();
//		} catch (Exception e) {
//			Log.warm(e.getCause().getMessage());
//		} 
		
//	exam OS type
		Properties pro = System.getProperties();
		String platform = pro.getProperty("os.name");
		if(platform.startsWith("Mac")||platform.startsWith("mac")){
			 OStype =1;
		}else if(platform.startsWith("Win")||platform.startsWith("win")){
			 OStype =2;
			 try {
				UIManager.setLookAndFeel("com.sun.java.swing.plaf.mac.MacLookAndFeel");
			} catch (ClassNotFoundException e) {
				Log.warm(e.getCause().getMessage());
			} catch (InstantiationException e) {
				Log.warm(e.getCause().getMessage());
			} catch (IllegalAccessException e) {
				Log.warm(e.getCause().getMessage());
			} catch (UnsupportedLookAndFeelException e) {
				Log.warm(e.getCause().getMessage());
			}
			 
		}
		new windows();              
	}
	

}
