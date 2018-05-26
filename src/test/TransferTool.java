package test;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TransferTool {

    public static void els2pdf(String inFilePath,String outFilePath){  
    	 ComThread.InitSTA(true);  
    	    ActiveXComponent ax=new ActiveXComponent("Excel.Application");  
    	    try{  
    	        ax.setProperty("Visible", new Variant(false));  
    	        ax.setProperty("AutomationSecurity", new Variant(3)); //���ú�  
    	        Dispatch excels=ax.getProperty("Workbooks").toDispatch();  
    	  
    	        Dispatch excel=Dispatch.invoke(excels,"Open",Dispatch.Method,new Object[]{  
    	            inFilePath,  
    	            new Variant(false),  
    	            new Variant(false)  
    	        },  
    	        new int[9]).toDispatch();  
    	        //ת����ʽ  
    	        Dispatch.invoke(excel,"ExportAsFixedFormat",Dispatch.Method,new Object[]{  
    	            new Variant(0), //PDF��ʽ=0  
    	            outFilePath,  
    	            new Variant(0)  //0=��׼ (���ɵ�PDFͼƬ�����ģ��) 1=��С�ļ� (���ɵ�PDFͼƬ����һ����Ϳ)  
    	        },new int[1]);  
    	  
    	        //�������ʹ��SaveAs  
    	        /*Dispatch.invoke(excel,"SaveAs",Dispatch.Method,new Object[]{ 
    	            outFile, 
    	            new Variant(57), 
    	            new Variant(false), 
    	            new Variant(57),  
    	            new Variant(57), 
    	            new Variant(false),  
    	            new Variant(true), 
    	            new Variant(57),  
    	            new Variant(true), 
    	            new Variant(true),  
    	            new Variant(true) 
    	        },new int[1]);*/  
    	  
    	        Dispatch.call(excel, "Close",new Variant(false));  
    	  
    	        if(ax!=null){  
    	            ax.invoke("Quit",new Variant[]{});  
    	            ax=null;  
    	        }  
    	        ComThread.Release();  
    	    }catch(Exception es){  
    	    }  
    }    
  
  
  
public static void main(String[] args) {  
    els2pdf("D:\\Excels\\ӡʵ�ɹ���ģ��880d496b-10f9-44d1-a538-78ee299bf8fb.xlsx","d:\\pdf.pdf");  
}  
}
