package test;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Properties;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class test {

	public static void main(String[] args) throws Exception {
		// 是否可预览标志 0普通文档，MP3，SWF，可以预览 ， 1 不支持的文件类型，不能预览 2，MP4播放
		String previewFlag = "0";
		
	}

	/**
	 * 清理缓冲区
	 * 
	 * @param isi
	 * @param ise
	 */
	public static void clearCache(InputStream isi, InputStream ise) {
		try {
			System.out.println("clearCache===========" + isi.toString() + "============" + ise.toString());
			final InputStream is1 = isi;
			// 启用单独线程清空InputStream缓冲区
			new Thread(new Runnable() {
				public void run() {
					System.out.println("before BufferedReader=========================");
					BufferedReader br = new BufferedReader(new InputStreamReader(is1));
					try {
						while (br.readLine() != null)
							;
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}).start();
			// 读入ErrorStream缓冲
			System.out.println("333333333333333333333333");
			BufferedReader br = new BufferedReader(new InputStreamReader(ise));
			// 保存缓冲输出结果
			StringBuilder buf = new StringBuilder();
			String line = null;
			try {
				System.out.println("before readLine==================");
				line = br.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			// 循环等待进程结束
			while (line != null)
				buf.append(line);
			is1.close();
			ise.close();
			br.close();
		} catch (Exception e) {
			System.out.println("clearCacheException===========" + e + "============");
			e.printStackTrace();
		}
	}

	/**
	 * 判断所转换文件类型是否合法
	 * 
	 * @param getFileType
	 *            //文件格式
	 * @param fileLegalFlag
	 *            //是否合法标志 false：非法 true：合法
	 */

	public static boolean isLegal(String getFileType) {
		boolean fileLegalFlag = false;
		if (getFileType.equals("TXT")) {
			fileLegalFlag = true;
		} else if (getFileType.equals("DOC") || getFileType.equals("DOCX")) {
			fileLegalFlag = true;
		} else if (getFileType.equals("PPT") || getFileType.equals("PPTX")) {
			fileLegalFlag = true;
		} else if (getFileType.equals("XLS") || getFileType.equals("XLSX")) {
			fileLegalFlag = true;
		}
		return fileLegalFlag;
	}
}
