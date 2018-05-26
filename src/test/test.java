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
		// �Ƿ��Ԥ����־ 0��ͨ�ĵ���MP3��SWF������Ԥ�� �� 1 ��֧�ֵ��ļ����ͣ�����Ԥ�� 2��MP4����
		String previewFlag = "0";
		String pathString = "C:\\Users\\Administrator\\Downloads\\";
		String newFileName = "�ɹ�����EXT002201810014.xlsx";
		String startupPath = "C:\\Program Files (x86)\\";
		Runtime rt = Runtime.getRuntime();
		rt.exec(startupPath+"OpenOffice 4\\program\\soffice.exe -headless -nologo -norestore -accept=socket,host=localhost,port=8100;urp;StarOffice.ServiceManager");
		/************************************** ����Ϊ�ļ�Ԥ�� *****************************************************/
		if (previewFlag.equals("0")) {
			// ��������Ϣ
			Properties props = System.getProperties();
			// ��ͬϵͳ�ָ���
			String separator = props.getProperty("file.separator");
			// �������洢���ļ����� 201212171709212889.doc
			String sysFileName = pathString + newFileName;
			// linux ��
			// ��һ����б��/home/apache-tomcat-6.0.33/webapps/jd/WebRoot/upload/knowledge/201212181318342606.ppt
			String fullPath = pathString + newFileName;
			// ��sysFileName����һ��,����Ѿ����ڣ�����ÿ������
			String sysFileName_temp = sysFileName.substring(0, sysFileName.lastIndexOf("."));

			File sourceFile; // ת��Դ�ļ�
			File pdfFile; // PDFý���ļ�
			// ��PDF��ʽ�ļ�����ʽ
			sourceFile = new File(fullPath);
			pdfFile = new File(separator + sysFileName_temp + ".pdf");
			rt = Runtime.getRuntime();
			if (!pdfFile.exists()) {
				// ��ȡ���Ӷ���
				OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
				// ȡ������
				connection.connect();
				// �����ļ���ʽת������
				DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
				// ʵ���ļ���ʽת��
				converter.convert(sourceFile, pdfFile);
				// ������ת����PDF�ļ�
				pdfFile.createNewFile();
				// �ͷ�����
				connection.disconnect();
			}
		}

	}

	/**
	 * ��������
	 * 
	 * @param isi
	 * @param ise
	 */
	public static void clearCache(InputStream isi, InputStream ise) {
		try {
			System.out.println("clearCache===========" + isi.toString() + "============" + ise.toString());
			final InputStream is1 = isi;
			// ���õ����߳����InputStream������
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
			// ����ErrorStream����
			System.out.println("333333333333333333333333");
			BufferedReader br = new BufferedReader(new InputStreamReader(ise));
			// ���滺��������
			StringBuilder buf = new StringBuilder();
			String line = null;
			try {
				System.out.println("before readLine==================");
				line = br.readLine();
			} catch (IOException e) {
				e.printStackTrace();
			}
			// ѭ���ȴ����̽���
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
	 * �ж���ת���ļ������Ƿ�Ϸ�
	 * 
	 * @param getFileType
	 *            //�ļ���ʽ
	 * @param fileLegalFlag
	 *            //�Ƿ�Ϸ���־ false���Ƿ� true���Ϸ�
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
