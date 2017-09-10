package mainpkg;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.*;

public class MainApp {

	public static void main(String[] args) {
		MainApp test=new MainApp();
		String targetpath="E:\\房产\\备案价\\autodownload4\\";
		ArrayList filenames=test.getFileNames();
		test.downLoadFiles(filenames,targetpath);
		test.getLocations(targetpath);
//		test.getLocation(targetpath+"2017-5-9_盛和雅颂花园2、3幢住房销售价格备案表.xls");
//		System.out.print("aaaa");
	}
    
	/**
	 * 功能说明：通过访问房价备案查询的主页获取所有房价备案信息的列表，在依次访问各个列表进入房价备案的详情页获取文件的下载连接地址
	 * 输入参数：
	 * 		remotePath：房价备案查询的主页http://dgdp.dg.gov.cn/publicfiles/business/htmlfiles/dgfg/s43520/index.htm
	 * 		reg：房价备案信息的列表行匹配正则表达式.*s43520.*
	 * 		pattern：访问房价备案的详情页的关键信息获取表达式s43520/(\\d+)/(\\d+)
	 * 		remotePath：房价备案详情页的访问连接http://dgdp.dg.gov.cn/publicfiles/business/htmlfiles/dgfg/s43520/
	 * 		res：下载文件地址行的匹配正则表达式.*(/publicfiles///business/htmlfiles/dgwjj/cmsmedia/document/|发布日期).*
	 * 		pattern：下载文件关键信息的获取表达式发布日期.*(\\d{4}-\\d+-\\d+).*\\s.*(doc\\d+.\\w+).*target=_blank>(.*.xls\\w?)
	 * 输出参数:
	 * 		res：房价备案详情页的访问关键信息列表
	 * 		res2：下载文件的关键信息列表
	 * 引用：
	 *		getFileName用于获取每一个页的关键信息
	 * @return
	 */
	public ArrayList getFileNames(){
	    //进入获取房价备案主页获取所有房价备案详情页的访问连接
		String remotePath="http://dgdp.dg.gov.cn/publicfiles/business/htmlfiles/dgfg/s43520/index.htm";
		ArrayList<String> res=new ArrayList<String>();
		res.clear();
		String reg=".*s43520.*";
		String pattern="s43520/(\\d+)/(\\d+)";
		getFileName(remotePath,res,reg,pattern);
		Iterator it=res.iterator();
		//依次进入方案备案详情页获取文件下载连接
		remotePath="http://dgdp.dg.gov.cn/publicfiles/business/htmlfiles/dgfg/s43520/";
		reg=".*(/publicfiles///business/htmlfiles/dgwjj/cmsmedia/document/|发布日期).*";
		pattern="发布日期.*(\\d{4}-\\d+-\\d+).*\\s.*(doc\\d+.\\w+).*target=_blank>(.*.xls\\w?)";
		ArrayList<String> res2=new ArrayList<String>();
		while(it.hasNext()){
			getFileName(remotePath+it.next()+"/"+it.next()+".htm",res2,reg,pattern);
		}
		return res2;
	}
	
	/**功能说明：下载给定列表中的所有文件
	 * @param filenames
	 */
	public void downLoadFiles(ArrayList filenames,String targetpath) {
		String remotePath="http://dgdp.dg.gov.cn/publicfiles///business/htmlfiles/dgwjj/cmsmedia/document/";
//		String remotePath = "http://dgdp.dg.gov.cn/publicfiles/business/htmlfiles/dgfg/s43520/201705/1127729.htm";
		String filename = "";
		String docname="";
		String localPath = targetpath;
        Iterator it=filenames.iterator(); 
		while(it.hasNext()) {
			filename=(String) it.next();
			docname=(String) it.next();
			filename=filename+"_"+it.next();
			downloadFile(remotePath + docname, localPath + filename);
			System.out.println(docname+"\t"+filename);
		}
	}

	/**功能说明：下载给定路径的单个文件
	 * @param remoteFilePath
	 * @param localFilePath
	 */
	public void downloadFile(String remoteFilePath, String localFilePath) {
		URL urlfile = null;
		HttpURLConnection httpUrl = null;
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		File f = new File(localFilePath);
		try {
			urlfile = new URL(remoteFilePath);
			httpUrl = (HttpURLConnection) urlfile.openConnection();
			httpUrl.connect();
			if (httpUrl.getContentType().equals("application/vnd.ms-excel")) {
				System.out.println("存在" + remoteFilePath);
			} else {
				// System.out.println("不存在"+remoteFilePath);
				return;
			}

			bis = new BufferedInputStream(httpUrl.getInputStream());
			bos = new BufferedOutputStream(new FileOutputStream(f));
			int len = 2048;
			byte[] b = new byte[len];
			while ((len = bis.read(b)) != -1) {
				bos.write(b, 0, len);
			}
			bos.flush();
			bis.close();
			httpUrl.disconnect();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (null != bis)
					bis.close();
				if (null != bos)
					bos.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**功能说明：获取给定路径页面的关键信息
	 * @param remoteFilePath
	 * @param res
	 * @param reg
	 * @param pattern
	 */
	public void getFileName(String remoteFilePath, ArrayList res,String reg,String  pattern ) {
		URL urlfile = null;
		HttpURLConnection httpUrl = null;
		BufferedReader bis = null;
		try {
			urlfile = new URL(remoteFilePath);
			httpUrl = (HttpURLConnection) urlfile.openConnection();
			httpUrl.connect();
			if (httpUrl.getContentType().equals("text/html; charset=utf-8")) {
			} else {
				 System.out.println("Error"+remoteFilePath);
				return;
			}
			bis = new BufferedReader(new InputStreamReader(httpUrl.getInputStream(), "utf-8"));
			String c=null;
			String str="";
			while ((c = bis.readLine()) != null) {
				if(c.matches(reg)){str=str+"\n"+c;}
				
			}
//			System.out.println(str);
			Pattern r = Pattern.compile(pattern);
	        Matcher m = r.matcher(str);
	        while(m.find())
	        for (int i=1;i<=m.groupCount();i++ ){
	        	res.add(m.group(i));
	        }  
	        
			bis.close();
			httpUrl.disconnect();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (null != bis)
					bis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void getLocation(String filepath) {

		File file = new File(filepath);
		File newfile=null;
		InputStream in = null;
		Workbook workbook = null;
		try {
			in = new FileInputStream(file);
			workbook = Workbook.getWorkbook(in);
			String location = "";
			Sheet[] sheets=workbook.getSheets();
			for(Sheet sh:sheets){
				for(Cell cell:sh.getRow(3)){
					if(cell.getContents().contains("所在镇街")){
						location=cell.getContents();
						break;
					}
				}
			}

			in.close();
			if(""==location)return;
			else{
				location=location.substring(location.lastIndexOf("：")+1);
				System.out.println(filepath.substring(0,filepath.lastIndexOf("\\")+1)+location+"_"+file.getName());
				newfile=new File(filepath.substring(0,filepath.lastIndexOf("\\")+1)+location+"_"+file.getName());
				file.renameTo(newfile);
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}catch(Exception e){
			e.printStackTrace();
		}
		finally {
			if(null!=workbook)workbook.close();
			try {
				in.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}



	/**
	 * @Function:根据表格中的位置信息修改文件名
	 * @param path
	 * @Datetime:2017-05-23 上午7:43:29
	 */
	public void getLocations(String path){
		File files=new File(path);
		String [] filenames=files.list();
		for(String file : filenames){
			System.out.print(file+"\n");
			getLocation(path+file);
		}
	}
	public void getUndownloadFilenames(String [] allfiles,String downloadedpath){
		
	}
}