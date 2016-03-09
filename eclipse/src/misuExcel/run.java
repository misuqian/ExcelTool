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
