package com.suite.lib.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.CRC32;
import java.util.zip.CheckedOutputStream;

import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipFile;
import org.apache.tools.zip.ZipOutputStream;

/**
 * 文件压缩控制器
 * @author Suite
 * 
 */
public class ZipCompressor {

	static final int BUFFER = 8192;
	
	/** 控制器实例的压缩文件 **/
	private File zipFile;
	
	/** 控制器实例的压缩文件的临时存储文件(文件夹) **/
	private File zipTempDirectory;
	
	/**
	 * 构造一个压缩文件控制器
	 * @param pathName -[压缩文件的路径]
	 */
	public ZipCompressor(String pathName) {
		zipFile = new File(pathName);
		zipTempDirectory = null;
	}
	
	/**
	 * 压缩文件
	 * @Description 生成的压缩文件会覆盖原有的压缩文件.
	 * @param srcPathName -[被压缩文件的路径]
	 * @return
	 */
	public boolean compress(String srcPathName) throws RuntimeException {
		return compress(srcPathName, false);
	}
	
	/**
	 * 压缩文件
	 * @param srcPathName -[被压缩文件的路径]
	 * @param isUpdate    -[是否更新已有的压缩文件]
	 * @return
	 */
	public boolean compress(String srcPathName, boolean isUpdate) throws RuntimeException {
		boolean result = false;
		/** 指定的被压缩文件 **/
		File file;
		if (null == srcPathName || !(file = new File(srcPathName)).exists()) {
			throw new RuntimeException(srcPathName + "不存在!");
		}
		/** 生成的压缩文件的输出流 **/
		ZipOutputStream zipOutputStream = null;
		try {
			//===== 读取已有的压缩文件数据,转存到临时文件(文件夹)之中 =====//
			/** 临时存储文件(文件夹)名称 **/
			String tempName;
			if (isUpdate) {
				//----- 需要保留已有压缩文件数据的情况 -----//
				try {
					//===== 生成已有压缩文件的临时存储文件 =====//
					tempName = zipFile.getName();
					tempName = new String( tempName.substring(0, tempName.lastIndexOf('.')) );
					// 当前服务器系统的时间戳
					long currentTimestamp = System.currentTimeMillis();
					/** 临时存储文件(文件夹)命名规则:[原文件不带后缀的名称]-[系统当前时间]-[随机数] **/
					tempName += "-"+ currentTimestamp +"-"+ Math.random();
					String tempPath = zipFile.getParent() +"\\"+ tempName;
					// 生成存放临时存储文件的临时文件夹
					zipTempDirectory = new File(tempPath);
					if (!zipTempDirectory.exists()) {
						zipTempDirectory.mkdirs();
					}
					// 解压压缩文件到临时文件夹之中
					decompress(tempPath);
					//=====//
				} catch (Exception e) {
					// 此时压缩文件不存在
				}
			}
			//===== 生成新的压缩文件 =====//
			FileOutputStream fileOutputStream = new FileOutputStream(zipFile);
			// 不加CRC32也能生成文件。此处数据校验机制未完全确认。
			CheckedOutputStream cos = new CheckedOutputStream(fileOutputStream, new CRC32());
			zipOutputStream = new ZipOutputStream(cos);
			zipOutputStream.setEncoding("GBK");
			String basedir = "";
			if (isUpdate && null != zipTempDirectory) {
				//----- 需要保留且存在已有压缩文件数据的情况 -----//
				try {
					//=== 装配临时文件数据到生成的压缩文件流中,并及时删除临时文件 ===//
					/** 临时文件夹中的文件集 **/
					File[] zipTempFiles;
					if (zipTempDirectory.exists() && zipTempDirectory.isDirectory() 
							&& null != (zipTempFiles = zipTempDirectory.listFiles())) {
						/**
						 * (JVM结束时)删除临时文件夹.
						 *
						 * 注明:此处声明是由于java.io.File.deleteOnExit()实现方法的执行顺序与声明顺序相反,
						 * 文件夹的删除声明需要在文件夹子目录的删除声明之前才可能执行成功.
						 */
						zipTempDirectory.deleteOnExit();
						// 遍历已有压缩文件中的文件集(临时文件夹中的文件集),逐一添加到新的压缩文件之中
						for (int i=0, i_maxsize = zipTempFiles.length; i < i_maxsize; i++) {
							File zipTempFile = zipTempFiles[i];
							compress(zipTempFile, zipOutputStream, basedir);
							try {
								if (!zipTempFile.delete()) {
									zipTempFile.deleteOnExit();
								}
							} catch (Exception e) {
								System.err.println("删除临时存储文件夹中的临时文件出错!错误信息如下:");
								e.printStackTrace();
								try {
									zipTempFile.deleteOnExit();
								} catch (Exception e1) {
									e1.printStackTrace();
								}
							}
						}
					}
				} catch (Exception e) {
					System.err.println("读取已有的压缩文件数据出错!错误信息如下:");
					e.printStackTrace();
				} finally {
					//=== 删除临时文件(文件夹) ===//
					if (null != zipTempDirectory && zipTempDirectory.exists()) {
						try {
							zipTempDirectory.delete();
						} catch (Exception e) {
							System.err.println("删除已有压缩文件的临时存储文件夹出错!错误信息如下:");
							e.printStackTrace();
						}
					}
				}
				//-----//
			}
			compress(file, zipOutputStream, basedir);
			//=====//
			result = true;
		} catch (Exception e) {
			RuntimeException e1 = new RuntimeException(e);
			e1.printStackTrace();
		} finally {
			//=== 关闭压缩文件的输出流 ===//
			if (null != zipOutputStream) {
				try {
					zipOutputStream.flush();
				} catch (Exception e) {
					e.printStackTrace();
				}
				try {
					zipOutputStream.close();
				} catch (Exception e) {
					System.err.println("压缩文件的输出流关闭失败!错误信息如下:");
					e.printStackTrace();
				}
			}
		}
		return result;
	}
	
	/**
	 * 压缩文件
	 * @Description !注意:此方法不关闭参数列表中的[压缩文件的输出流]<code>out</code>.
	 * @param file    -[被压缩的文件]
	 * @param out     -[压缩文件的输出流]
	 * @param basedir -[压缩文件的根目录]
	 * @return
	 */
	private boolean compress(File file, ZipOutputStream out, String basedir) {
		boolean result;
		//--- 判断文件目录或文件,并分类处理 ---//
		if (file.isDirectory()) {
			// 压缩文件目录
			result = this.compressDirectory(file, out, basedir);
		}else{
			// 压缩一个文件
			result = this.compressFile(file, out, basedir);
		}
		return result;
	}
	
	/**
	 * 压缩文件目录
	 * @Description 空文件夹不进行压缩处理.
	 * @param directory -[被压缩的文件目录]
	 * @param out       -[压缩文件的输出流]
	 * @param basedir   -[压缩文件的根目录]
	 * @return
	 */
	private boolean compressDirectory(File directory, ZipOutputStream out, String basedir) {
		if (null == directory || !directory.exists()) {
			return false;
		}
		boolean result = false;
		File[] files = directory.listFiles();
		if (null != files) {
			result = true;
			for (int i=0, i_maxsize = files.length; i < i_maxsize; i++) {
				//--- 递归 ---//
				if (!compress(files[i], out, basedir + directory.getName() + "/")) {
					result = false;
				}
			}
		}
		return result;
	}
	
	/**
	 * 压缩一个文件
	 * @Description !注意:此方法不关闭参数列表中的[压缩文件的输出流]<code>out</code>.
	 * @param file    -[被压缩的文件]
	 * @param out     -[压缩文件的输出流]
	 * @param basedir -[压缩文件的根目录]
	 */
	private boolean compressFile(File file, ZipOutputStream out, String basedir) {
		if (null == file || !file.exists()) {
			return false;
		}
		/** 被压缩文件的输入流 **/
		BufferedInputStream bis = null;
		/** 被压缩文件对应的压缩文件条目 **/
		ZipEntry entry = null;
		try {
			bis = new BufferedInputStream(new FileInputStream(file));
			// 创建压缩文件的条目
			entry = new ZipEntry(basedir + file.getName());
			// 将条目数据写入压缩文件输出流
			out.putNextEntry(entry);
			int count;
			byte[] data = new byte[BUFFER];
			while ((count = bis.read(data, 0, BUFFER)) != -1) {
				out.write(data, 0, count);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			// 关闭压缩文件的条目
			if (null != entry) {
				try {
					out.closeEntry();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			// 关闭被压缩文件的输入流
			if (null != bis) {
				try {
					bis.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		return true;
	}
	
	/**
	 * 解压文件
	 * @Description 被解压的压缩文件为当前控制器的压缩文件
	 * @param outputDirectoryPath -[从压缩文件中解压的文件集的输出路径]
	 * @return
	 */
	public boolean decompress(String outputDirectoryPath) throws RuntimeException {
		if (null == outputDirectoryPath) {
			throw new RuntimeException(outputDirectoryPath + "不存在!");
		}
		if (null == zipFile || !zipFile.exists()) {
			throw new RuntimeException("压缩文件不存在,无法解压!");
		}
		boolean result = false;
		/** 被解压的压缩文件 **/
		ZipFile zip = null;
		try {
			//===== 解压文件 =====//
			zip = new ZipFile(zipFile, "GBK");
			Enumeration<ZipEntry> entries = zip.getEntries();
			while (entries.hasMoreElements()) {
				/** 当前(遍历的)压缩文件条目 **/
				ZipEntry zipEntry = entries.nextElement();
				/** 当前条目的输入流 **/
				InputStream zipEntry_InputStream = null;
				/** (当前条目对应的)输出文件 **/
				File outputFile;
				/** (当前条目对应的)输出文件的路径 **/
				String outputPath;
				/** (当前条目对应的)输出文件的输出流 **/
				FileOutputStream outputFileOutputStream = null;
				//=== 初始化输出文件信息 ===//
				outputPath = outputDirectoryPath +"\\"+ zipEntry.getName();
				outputFile = new File(outputPath);
				//=== 读取当前条目数据输出到输出文件中 ===//
				if (zipEntry.isDirectory()) {
					//--- 当前条目为文件目录的情况 ---//
					outputFile.mkdirs();
				}else{
					//--- 当前条目为一个文件的情况 ---//
					File parent = outputFile.getParentFile();
					if (!parent.exists()) {
						parent.mkdirs();
					}
					try {
						zipEntry_InputStream = zip.getInputStream(zipEntry);
						outputFileOutputStream = new FileOutputStream(outputFile);
						int len;
						byte[] buff = new byte[1024];
						while ((len = zipEntry_InputStream.read(buff)) != -1) {
							outputFileOutputStream.write(buff, 0, len);
						}
					} catch (Exception e) {
						System.err.println("读取当前条目数据输出到输出文件出错!错误信息如下:");
						e.printStackTrace();
					} finally {
						//=== 关闭(当前条目对应的)输出文件的输出流 ===//
						if (null != outputFileOutputStream) {
							try {
								outputFileOutputStream.flush();
							} catch (Exception e) {
								e.printStackTrace();
							}
							try {
								outputFileOutputStream.close();
							} catch (Exception e) {
								e.printStackTrace();
							}
						}
						//=== 关闭当前文件的输入流 ===//
						if (null != zipEntry_InputStream) {
							zipEntry_InputStream.close();
						}
					}
				}
				//===//
			}
			//=====//
			result = true;
		} catch (Exception e) {
			RuntimeException e1 = new RuntimeException(e);
			e1.printStackTrace();
		} finally {
			// 关闭被解压的压缩文件
			try {
				zip.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return result;
	}
	
	/*public static void main(String[] args) {
		ZipCompressor zc = new ZipCompressor("D://Fax//IAEA资料信息//测试压缩.zip");
		boolean isSuccess = (zc.compress("D://Fax//历史案例表_c9f95a6f-d24d-457a-a2c5-33afa2f60fcd.xls", true)
				& zc.compress("D://Fax//历史案例表_ccc7806c-4ec7-45fc-b14b-ccee9fa7eeea.xls", true)
				& zc.compress("D://Fax//历史案例表.xls_18f7147a-cf2a-4a60-ad21-a59db1ae5bfb.xls", true)
				& zc.compress("D://Fax//测试---压缩文件", true)
		);
		System.err.println(isSuccess);
	}*/
	
}
