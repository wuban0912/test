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
		String pathString = "C:\\Users\\Administrator\\Downloads\\";
		String newFileName = "采购订单EXT002201810014.xlsx";
		String startupPath = "C:\\Program Files (x86)\\";
		Runtime rt = Runtime.getRuntime();
		rt.exec(startupPath+"OpenOffice 4\\program\\soffice.exe -headless -nologo -norestore -accept=socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
		/************************************** 以下为文件预览 *****************************************************/
		if (previewFlag.equals("0")) {
			// 服务器信息
			Properties props = System.getProperties();
			// 不同系统分隔符
			String separator = props.getProperty("file.separator");
			// 服务器存储的文件名称 201212171709212889.doc
			String sysFileName = pathString + newFileName;
			// linux 下
			// 是一个正斜杠/home/apache-tomcat-6.0.33/webapps/jd/WebRoot/upload/knowledge/201212181318342606.ppt
			String fullPath = pathString + newFileName;
			// 与sysFileName保持一致,如果已经存在，不再每次生成
			String sysFileName_temp = sysFileName.substring(0, sysFileName.lastIndexOf("."));

			File sourceFile; // 转换源文件
			File pdfFile; // PDF媒介文件
			// 非PDF格式文件处理方式
			sourceFile = new File(fullPath);
			pdfFile = new File(separator + sysFileName_temp + ".pdf");
			rt = Runtime.getRuntime();
			if (!pdfFile.exists()) {
				// 获取连接对象
				OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
				// 取得连接
				connection.connect();
				// 创建文件格式转换对象
				DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
				// 实现文件格式转换
				converter.convert(sourceFile, pdfFile);
				// 生成已转换的PDF文件
				pdfFile.createNewFile();
				// 释放连接
				connection.disconnect();
			}
		}

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
