package ntpc.ccai.servlet;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import jxl.Sheet;
import ntpc.bean.DBCon;
import ntpc.ccai.bean.StuBasis;
import ntpc.ccai.bean.StuBasisDataList;
import ntpc.ccai.bean.StuClass;
import ntpc.ccai.bean.StuRegister;
import ntpc.ccai.bean.UserData;
import ntpc.ccai.util.FileUtil;
import ntpc.ccai.util.ParseXLSUtil;
import ntpc.ccai.util.SystemUtil;
/**
 * Author : Josh, Daphne, Kimberly
 * Date   : 20170608
 * Purpose: 檔案上傳
 */
@WebServlet("/BTSchRoll")
public class BTSchRoll extends HttpServlet {
    private static final long serialVersionUID = 1L;
    private Logger      logger              = LogManager.getLogger(this.getClass());
    private String      INIT                = "/WEB-INF/jsp/BTSchRoll.jsp";         // 在校生新增異動頁面
    private String      ERROR               = "/WEB-INF/jsp/ErrorPage.jsp";         // 失敗的 view
    
    protected void service(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        String error_msg = ""; // 錯誤訊息
        String targetURL = "";
        
        // 收集資料
        request.setCharacterEncoding("UTF-8");      // Ajax need
        response.setContentType("text/html;charset=UTF-8");
        HttpSession session = request.getSession(); // 取得 session 物件
        
        // 設定資料
        UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud

        // Steps 1 : 權限檢查
        if (ud == null) {
            // 尚未登入
            error_msg = "連線逾期，無法順利建立資料，請重新登入";
            request.setAttribute("error_msg", error_msg);
            targetURL = ERROR;
            SystemUtil.forward(targetURL, request, response);
            return;
        }
        else {
//        	System.out.println("getRole_code=== >" +ud.getRole_code());
			if (!ud.getRole_code().equals("sta")) { // sta承辦人
				// 沒有使用權限
				error_msg = "<span style='color:red;'>您沒有使用權限</span>";
				request.setAttribute("error_msg", error_msg); // 錯誤訊息
				targetURL = ERROR;
				SystemUtil.forward(targetURL, request, response);
				return;
			}
		}
        
        // Step 2: doGet or doPost
        super.service(request, response);
    }
    
    protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doPost(request, response);
    }

    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        String            targetURL = "";        // forward view
        String         deleteResult = "";
        DBCon                   dbc = null;
        Connection             conn = null;
        StuClass           stuClass = new StuClass();
        StuBasis           stuBasis = new StuBasis();
        StuRegister          stuReg = new StuRegister();
        List<StuClass>    classList = null;
        List<StuBasis>    basisList = null;
        List<StuRegister>   regList = null;
        String           uploadFlag = "";       //上傳的動作的旗標
        String           fileaction = "";       //下載的table
        
        try{
            targetURL = this.INIT;
            
            // 收集資料
            request.setCharacterEncoding("UTF-8");      // Ajax need
            HttpSession session = request.getSession(); // 取得 session 物件
            
            // 設定資料
            UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud
            String Sch_code = ud.getSch_code();         // 取得學校代碼
            String adg_code = ud.getAdg_code();         // 取得學校部別代碼
            dbc  = new DBCon(Sch_code);
            conn = dbc.getConnection();
            
            // 上傳功能
            uploadFlag = request.getParameter("uploadFlag");
            //logger.debug("uploadFlag:" + uploadFlag);
            if("Y".equals(uploadFlag)){
//            	logger.info("uploadFlag == Y");
                this.doPostUpload(request, response);
            }
            
            // 下載功能
            fileaction = request.getParameter("fileaction");    //取得頁面的action
            classList = stuClass.getSBJYear(conn);
            basisList = stuBasis.getYear(conn);
            regList   = stuReg.getSbjYear(conn);
            request.setAttribute("classList", classList);
            request.setAttribute("basisList", basisList);
            request.setAttribute("regList", regList);
            
            if("downloadClass".equals(fileaction))
            {
                String sbjYear = request.getParameter("sbjYear");
                String sbjSem = request.getParameter("sbjSem");
                this.doPostExportClassExcel(request, response, sbjYear, sbjSem);
            }
            else if("downloadBasis".equals(fileaction))
            {
                String cmat_year= request.getParameter("cmat_year");
                this.doPostExportBasisExcel(request, response, cmat_year);
            }
            else if("downloadRegister".equals(fileaction))
            {
                String sbjYear = request.getParameter("sbjYear");
                String sbjSem = request.getParameter("sbjSem");
                this.doPostExportRegisterExcel(request, response, sbjYear, sbjSem);
            }
            
            // 刪除功能
            String action = request.getParameter("action");
            //ParseXLSUtil excel = new ParseXLSUtil();
            
            if("deleteClass".equals(action))
            {
                String sbjYear = request.getParameter("sbjYear");
                String sbjSem = request.getParameter("sbjSem");
                
                // 1.先下載備份檔案
                //this.doPostExportClassExcel(request, response, sbjYear, sbjSem);
                // 2.刪除檔案
                deleteResult = stuClass.delClassAll(conn, sbjYear, sbjSem, adg_code);
                request.setAttribute("delMsg",deleteResult);
                if(deleteResult.contentEquals("資料刪除成功！")) {
                    logger.info("[school_roll] BTSchRoll doPost deleteClass success! " + ntpc.util.StringUtil.getLogInfo(request) + " sbjYear=" + sbjYear + " sbjSem=" + sbjSem);
                }
            } 
            else if("deleteBasis".equals(action))
            {
                String cmat_year= request.getParameter("cmat_year");
                
                // 1.先下載備份檔案
                //this.doPostExportBasisExcel(request, response, cmat_year);
                
                // 2.刪除stu_id_data資料    1061101 (Daphne)
                basisList = stuBasis.queryRgnoOfYear(conn, cmat_year, adg_code);
                for(int i=0; i<basisList.size(); i++){
                    StuBasis sb = new StuBasis();
                    sb = basisList.get(i);
                    int rgno = sb.getRgno();
                    deleteResult = stuBasis.deleteStuIdData(conn, rgno);
                }
                
                // 3.刪除stu_basis資料
                deleteResult = stuBasis.deleteAll(conn,cmat_year, adg_code);
                request.setAttribute("delMsg",deleteResult);
                if(deleteResult.contentEquals("資料刪除成功！")) {
                    logger.info("[school_roll] BTSchRoll doPost deleteBasis success! " + ntpc.util.StringUtil.getLogInfo(request) + " cmat_year=" + cmat_year);
                }
            }
            else if("deleteRegister".equals(action)) 
            {
                String sbjYear = request.getParameter("sbjYear");
                String sbjSem = request.getParameter("sbjSem");
                // 1.先下載備份檔案
                //this.doPostExportRegisterExcel(request, response, sbjYear, sbjSem);
                // 2.刪除檔案
                deleteResult = stuReg.delRegisterAll(conn, sbjYear, sbjSem, adg_code);
                request.setAttribute("delMsg",deleteResult);
                if(deleteResult.contentEquals("資料刪除成功！")) {
                    logger.info("[school_roll] BTSchRoll doPost deleteRegister success! " + ntpc.util.StringUtil.getLogInfo(request) + " sbjYear=" + sbjYear + " sbjSem=" + sbjSem);
                }
            }
        }catch (Exception e) {
            logger.error("查詢學校資料列表發生錯誤!", e);
            request.setAttribute("error_msg", "查詢學校資料列表發生錯誤!");
            targetURL = ERROR;
            e.printStackTrace();
        }finally{
            dbc.closeCon();
        }
        request.setCharacterEncoding("UTF-8");
        SystemUtil.forward(targetURL, request, response);
    }
    
    
    /**Purpose: 將excel資料匯入資料庫
     * Author : Kimberly
     * Date   : 20170802
     * Revise : 20180521 | Josh | 調整insert和update的StuIdData寫入時序
     */
    @SuppressWarnings("unchecked")
    protected void doPostUpload(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        String method    = "";        // 操作動作(新增new或修改revise)
        String table     = "";        // 操作table(class或basis或register)
        String Msg       = "";        // 訊息
//        String targetURL = this.INIT;;  
        
        File   tmpDir = null;
        List<FileItem> items = null;  // 接收從view端傳來的form資料或檔案
        
        // 收集資料
        request.setCharacterEncoding("UTF-8");      // Ajax need
        HttpSession session = request.getSession(); // 取得 session 物件
        
        // 設定資料
        UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud
        String Sch_code = ud.getSch_code();         // 取得學校代碼
        String adg_code = ud.getAdg_code();         // 取得學校部別代碼
        DBCon dbc  = new DBCon(Sch_code);
        Connection conn = null;
        
      
        try {
            conn = dbc.getConnection();
        
            // 取得要上傳的方式與Table
            method = request.getParameter("method"); 
            table  = request.getParameter("table");  
            
            logger.debug("method:" + method);
            logger.debug("table:" + table);
        
            // 使用POI拆解excel
            ParseXLSUtil<?> excel = new ParseXLSUtil<Object>();
            
            FileItem ExcelItem = null;
            // Create a factory for disk-based file items
            DiskFileItemFactory factory = new DiskFileItemFactory();
            
            // Set factory constraints
            factory.setSizeThreshold(1024*1024);  // 限制緩衝區大小
            factory.setRepository(tmpDir);        // 檔案超過threshold大小的的暫存目錄
            //factory.setSizeMax(1024*1024*4);    // 限制檔案大小
            
            // Create a new file upload handler
            ServletFileUpload fileUpload = new ServletFileUpload(factory);
            
            // Parse the request
            items = (List<FileItem>)fileUpload.parseRequest(request);
            // Process the uploaded items
            Iterator<FileItem> iter = items.iterator();
            while (iter.hasNext()) {
                FileItem item = (FileItem) iter.next();
                if (!item.isFormField()) {
                    // 檢查副檔名
                    Msg = FileUtil.checkFileExtName(item.getName(), ".xls");
                    ExcelItem = item;
                }
                logger.debug("ExcelItem:" + ExcelItem);
            }
            
            if (!"".equals(Msg)){
                request.setAttribute("msg", Msg);
            }else{
                
                //上傳班級代碼表
                if ("class".equals(table)) {
                    ArrayList<StuClass> list = (ArrayList<StuClass>) excel.parseData(ExcelItem, table);
//                    logger.debug("list.size():" + list.size());
                    StuClass stuClass = new StuClass();
                    String divCode = stuClass.getDivCode(conn);
                    if ("new".equals(method)) {
                        for (int i = 0; i < list.size(); i++) {
                            boolean is_right = stuClass.checkDepartment(conn, list.get(i).getDIV_CODE(), list.get(i).getEDU_DEP_CODE());
//                            logger.debug("is_right:" + is_right);
                            int result = 0 ;
                            if (!is_right) {
                            	result = 1;
                            	Msg += "第 " + (i + 1) + " 筆「班群代碼」和「教育部科別代碼」不符合，\\n";                            	
                            }
                            if (Sch_code.equals("040B02") || Sch_code.equals("074B23") || Sch_code.equals("183B07") || Sch_code.equals("200B03") || Sch_code.equals("210B05")) {
                            	
                            } else {
                            	if (list.get(i).getEDU_DEP_CODE().equals("101") && list.get(i).getDIV_CODE().equals(divCode)) {
                            		if (result == 0) {
                            			Msg += "第 " + (i + 1) + " 筆「教育部科別代碼」如果為普通科101，「班群代碼」不適用代碼 " + divCode + "，\\n";
                            			result = 1;                            		
                            		} else {
                            			Msg += "「教育部科別代碼」如果為普通科101，「班群代碼」不適用代碼 " + divCode + "，\\n";
                            		}
                            	}
                            }
                        }
                        if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                        if (Msg.equals("")) {
                            Msg = stuClass.insertClass(conn, list); // 20170623
                            logger.info("[school_roll] BTSchRoll doPostUpload class new success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                        }
                        request.setAttribute("msg", Msg);
						
                    } else if ("revise".equals(method)) {
                        for (int i = 0; i < list.size(); i++) {
                            boolean is_right = stuClass.checkDepartment(conn, list.get(i).getDIV_CODE(), list.get(i).getEDU_DEP_CODE());
//                            logger.debug("is_right:" + is_right);
                            int result = 0 ;
                            if (!is_right) {
                            	result = 1;
                            	Msg += "第 " + (i + 1) + " 筆「班群代碼」和「教育部科別代碼」不符合，\\n";                            	
                            }
                            if (Sch_code.equals("040B02") || Sch_code.equals("074B23") || Sch_code.equals("183B07") || Sch_code.equals("200B03") || Sch_code.equals("210B05")) {
                            	
                            } else {
                            	if (list.get(i).getEDU_DEP_CODE().equals("101") && list.get(i).getDIV_CODE().equals(divCode)) {
                            		if (result == 0) {
                            			Msg += "第 " + (i + 1) + " 筆「教育部科別代碼」如果為普通科101，「班群代碼」不適用代碼 " + divCode + "，\\n";
                            			result = 1;                            		
                            		} else {
                            			Msg += "「教育部科別代碼」如果為普通科101，「班群代碼」不適用代碼 " + divCode + "，\\n";
                            		}
                            	}
                            }
                        }
                        if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                        if (Msg.equals("")) {
                            Msg = stuClass.updateClass(conn, list);
                            logger.info("[school_roll] BTSchRoll doPostUpload class revise success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                        }
                        request.setAttribute("msg", Msg);
                    }
                }
                
                //上傳學生基本資料
                else if("basis".equals(table)){
                    ArrayList<StuBasis> list = (ArrayList<StuBasis>) excel.parseData(ExcelItem,table);
                    logger.debug(" getErrorcolumn ：" + list.get(0).getErrorcolumn());
            		logger.debug(" getErrorrow ：" + list.get(0).getErrorrow());
                    if(list.get(0).getErrorcolumn() !=0 && list.get(0).getErrorrow() !=0) {
                    	logger.debug(" getErrorcolumn ：" + list.get(0).getErrorcolumn());
                    	logger.debug(" getErrorrow ：" + list.get(0).getErrorrow());
                		Msg = "第"+ list.get(0).getErrorcolumn() + "直欄 ， 第" + list.get(0).getErrorrow().toString() + "橫列錯誤。\\n -------------------------\\n 不符合Excel匯入規定格式請參考 。 \\n 5.欄位說明表[學生基本資料]" ; 
                		request.setAttribute("msg", Msg);
                		dbc.closeCon();
                		return;
                    }         
                    
                    StuBasis stuBasis = new StuBasis();
                    if("new".equals(method)){
                    	for (int i = 0; i < list.size(); i++) {
                    		stuBasis = list.get(i);
                    		boolean idno = stuBasis.checkIdno(conn, adg_code, stuBasis.getIdno(), stuBasis.getCmat_year());
                    		boolean schCode = stuBasis.checkSchCode(conn, stuBasis.getSch_code());
                    		boolean perCityCode = stuBasis.checkCityCode(conn, stuBasis.getPer_city_code());
                    		boolean comCityCode = stuBasis.checkCityCode(conn, stuBasis.getCom_city_code());
                    		boolean perCityTown = stuBasis.checkCityTown(conn, stuBasis.getPer_city_code(), stuBasis.getPer_town_num());
                    		boolean comCityTown = stuBasis.checkCityTown(conn, stuBasis.getCom_city_code(), stuBasis.getCom_town_num());
                    		boolean perCityTownVil = stuBasis.checkCityTownVil(conn, stuBasis.getPer_city_code(), stuBasis.getPer_town_num(), stuBasis.getPer_village_num());
                    		boolean comCityTownVil = stuBasis.checkCityTownVil(conn, stuBasis.getCom_city_code(), stuBasis.getCom_town_num(), stuBasis.getCom_village_num());
                    		boolean per_isNumeric = stuBasis.getPer_neighbor().matches("[+-]?\\d*(\\.\\d+)?"); //戶籍鄰
                    		boolean com_isNumeric = stuBasis.getCom_neighbor().matches("[+-]?\\d*(\\.\\d+)?"); //通訊鄰
                    		
                    		int result = 0 ;
                    		if(idno) {
                    			Msg += "第 " + (i + 1) + " 筆「基本資料」已上傳過，基本資料僅需入學時上傳一次即可，\\n";
                    			result = 1;
                    		}
                    		if(!schCode) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「入學資格學校代碼」不符合學校代碼格式，\\n";
                    				result = 1;                    				
                    			} else {
                    				Msg += "「入學資格學校代碼」不符合學校代碼格式，\\n";
                    			}
                    		}
                    		if("".equals(stuBasis.getId_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「報部身份」不符合身份代碼格式，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「報部身份」不符合身份代碼格式，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_city_code()) && "".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_village_num()) && !perCityCode) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」不符合地址縣市名稱，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」不符合地址縣市名稱，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_village_num()) && !perCityTown) { // && 比對縣市&鄉鎮(如果有寫縣市和鄉鎮)
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 三欄都有寫就比對
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_town_num()) && !"".equals(stuBasis.getPer_village_num()) && !perCityTownVil) { // && 比對縣市&鄉鎮&村里
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」、「戶籍地址村里」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」、「戶籍地址村里」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫鄉鎮，縣市不可為空
                    		if(!"".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_city_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫村里，鄉鎮不可為空
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_village_num()) && "".equals(stuBasis.getPer_town_num())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_neighbor()) && !per_isNumeric) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址鄰」請填數字，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址鄰」請填數字，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_city_code()) && "".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_village_num()) && !comCityCode) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」不符合地址縣市名稱，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」不符合地址縣市名稱，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_village_num())  && !comCityTown) { // && 比對縣市&鄉鎮(如果有寫縣市和鄉鎮)
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 三欄都有寫就比對
                    		if(!"".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_town_num()) && !"".equals(stuBasis.getCom_village_num()) && !comCityTownVil) { // && 比對縣市&鄉鎮&村里
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」、「通訊地址村里」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」、「通訊地址村里」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫鄉鎮，縣市不可為空
                    		if(!"".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_city_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫村里，鄉鎮不可為空
                    		if(!"".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_village_num()) && "".equals(stuBasis.getCom_town_num())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_neighbor()) && !com_isNumeric) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址鄰」請填數字，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址鄰」請填數字，\\n";
                    			}
                    		}
                    	}
                    	if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                    	if (Msg.equals("")) {
                    		Msg = stuBasis.insertBasis(conn, list);     //20170623
                    		logger.info("[school_roll] BTSchRoll doPostUpload basis new success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                    	}
//                        stuBasis.insertStuIdData(conn, list); //20171101 (Daphne)  20180521 (Josh轉為insertBasis時進行insertStuIdData)
                        request.setAttribute("msg", Msg);
                    }
                    else if("revise".equals(method)){
                    	for (int i = 0; i < list.size(); i++) {
                    		stuBasis = list.get(i);
                    		boolean schCode = stuBasis.checkSchCode(conn, stuBasis.getSch_code());
                    		boolean perCityCode = stuBasis.checkCityCode(conn, stuBasis.getPer_city_code());
                    		boolean comCityCode = stuBasis.checkCityCode(conn, stuBasis.getCom_city_code());
                    		boolean perCityTown = stuBasis.checkCityTown(conn, stuBasis.getPer_city_code(), stuBasis.getPer_town_num());
                    		boolean comCityTown = stuBasis.checkCityTown(conn, stuBasis.getCom_city_code(), stuBasis.getCom_town_num());
                    		boolean perCityTownVil = stuBasis.checkCityTownVil(conn, stuBasis.getPer_city_code(), stuBasis.getPer_town_num(), stuBasis.getPer_village_num());
                    		boolean comCityTownVil = stuBasis.checkCityTownVil(conn, stuBasis.getCom_city_code(), stuBasis.getCom_town_num(), stuBasis.getCom_village_num());
                    		boolean per_isNumeric = stuBasis.getPer_neighbor().matches("[+-]?\\d*(\\.\\d+)?"); //戶籍鄰
                    		boolean com_isNumeric = stuBasis.getCom_neighbor().matches("[+-]?\\d*(\\.\\d+)?"); //通訊鄰
                    		
                    		int result = 0 ;
                    		if(!schCode) {
                    			Msg += "第 " + (i + 1) + " 筆「入學資格學校代碼」不符合學校代碼格式，\\n";
                    			result = 1;
                    		}
                    		if("".equals(stuBasis.getId_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「報部身份」不符合身份代碼格式，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「報部身份」不符合身份代碼格式，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_city_code()) && "".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_village_num()) && !perCityCode) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」不符合地址縣市名稱，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」不符合地址縣市名稱，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_village_num()) && !perCityTown) { // && 比對縣市&鄉鎮(如果有寫縣市和鄉鎮)
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 三欄都有寫就比對
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_town_num()) && !"".equals(stuBasis.getPer_village_num()) && !perCityTownVil) { // && 比對縣市&鄉鎮&村里
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」、「戶籍地址村里」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」、「戶籍地址村里」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫鄉鎮，縣市不可為空
                    		if(!"".equals(stuBasis.getPer_town_num()) && "".equals(stuBasis.getPer_city_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址縣市」、「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫村里，鄉鎮不可為空
                    		if(!"".equals(stuBasis.getPer_city_code()) && !"".equals(stuBasis.getPer_village_num()) && "".equals(stuBasis.getPer_town_num())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getPer_neighbor()) && !per_isNumeric) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「戶籍地址鄰」請填數字，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「戶籍地址鄰」請填數字，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_city_code()) && "".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_village_num()) && !comCityCode) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」不符合地址縣市名稱，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」不符合地址縣市名稱，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_village_num())  && !comCityTown) { // && 比對縣市&鄉鎮(如果有寫縣市和鄉鎮)
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 三欄都有寫就比對
                    		if(!"".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_town_num()) && !"".equals(stuBasis.getCom_village_num()) && !comCityTownVil) { // && 比對縣市&鄉鎮&村里
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」、「通訊地址村里」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」、「通訊地址村里」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫鄉鎮，縣市不可為空
                    		if(!"".equals(stuBasis.getCom_town_num()) && "".equals(stuBasis.getCom_city_code())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址縣市」、「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		// 20220328 有寫村里，鄉鎮不可為空
                    		if(! "".equals(stuBasis.getCom_city_code()) && !"".equals(stuBasis.getCom_village_num()) && "".equals(stuBasis.getCom_town_num())) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址鄉鎮」不符合地址相關資料，\\n";
                    			}
                    		}
                    		if(!"".equals(stuBasis.getCom_neighbor()) && !com_isNumeric) {
                    			if (result == 0) {
                    				Msg += "第 " + (i + 1) + " 筆「通訊地址鄰」請填數字，\\n";
                    				result = 1;
                    			} else {
                    				Msg += "「通訊地址鄰」請填數字，\\n";
                    			}
                    		}                    		
                    	}
                    	if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                    	if (Msg.equals("")) {
                    		Msg = stuBasis.updateBasis(conn, list);
                    		logger.info("[school_roll] BTSchRoll doPostUpload basis revise success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                    	}
//                        stuBasis.updateStuIdData(conn, list); //20171101 (Daphne)  20180521 (Josh轉為updateBasis時進行updateStuIdData)
                        request.setAttribute("msg", Msg);
                    }

                }
                
                //上傳註冊編班資料
                else if("register".equals(table)){
                    ArrayList<StuRegister> list = (ArrayList<StuRegister>) excel.parseData(ExcelItem,table);
                    StuRegister stuRegister = new StuRegister();
                    String divCode = stuRegister.getDivCode(conn);
                    if("new".equals(method)){
                    	for (int i = 0; i < list.size(); i++) {
                    		int result = 0 ;
                    		if("".equals(list.get(i).getCLS_NO())) {
                    			Msg += "第 " + (i + 1) + " 筆「班級座號」為必填欄位，\\n";
                    			result = 1;
                    		}
                    		if (Sch_code.equals("040B02") || Sch_code.equals("074B23") || Sch_code.equals("183B07") || Sch_code.equals("200B03") || Sch_code.equals("210B05")) {
                    			
                    		} else {
                    			if (!"".equals(divCode) && list.get(i).getDIV_CODE().equals(divCode)) {
                    				if (result == 0) {
                    					Msg += "第 " + (i + 1) + " 筆「班群代碼」的科別如果屬於普通科101，則不適用代碼 " + divCode + "，\\n";
                    					result = 1;
                    				} else {
                    					Msg += "「班群代碼」的科別如果屬於普通科101，則不適用代碼 " + divCode + "，\\n";
                    				}
                    			}
                    		}
                    	}
                    	if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                    	if (Msg.equals("")) {
                            Msg = stuRegister.insertReg(conn, list);   //20170623
                            logger.info("[school_roll] BTSchRoll doPostUpload register new success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                    	}
                        request.setAttribute("msg", Msg);
                    }
                    else if("revise".equals(method)){
                    	for (int i = 0; i < list.size(); i++) {
                    		int result = 0 ;
                    		if("".equals(list.get(i).getCLS_NO())) {
                    			Msg += "第 " + (i + 1) + " 筆「班級座號」為必填欄位，\\n";
                    			result = 1;
                    		}
                    		if (Sch_code.equals("040B02") || Sch_code.equals("074B23") || Sch_code.equals("183B07") || Sch_code.equals("200B03") || Sch_code.equals("210B05")) {
                    			
                    		} else {
                    			if (!"".equals(divCode) && list.get(i).getDIV_CODE().equals(divCode)) {
                    				if (result == 0) {
                    					Msg += "第 " + (i + 1) + " 筆「班群代碼」的科別如果屬於普通科101，則不適用代碼 " + divCode + "，\\n";
                    					result = 1;
                    				} else {
                    					Msg += "「班群代碼」的科別如果屬於普通科101，則不適用代碼 " + divCode + "，\\n";
                    				}
                    			}
                    		}
                    	}
                    	if (!"".equals(Msg)) {
                    		Msg += "以上請確認！\\n";
                    	}
                    	if (Msg.equals("")) {
                            Msg = stuRegister.updateReg(conn, list);   //20170623
                            logger.info("[school_roll] BTSchRoll doPostUpload register revise success! " + ntpc.util.StringUtil.getLogInfo(request) + " list size=" + list.size());
                    	}
                        request.setAttribute("msg", Msg);
                    }                       
                }
                request.setAttribute("message", Msg);
            }       
        } catch (Exception e) {
            logger.error("匯入資料錯誤，請聯絡系統管理人員!", e);
            Msg = "1.無法辨識、取得上傳Excel檔案資料，請檢查Excel是否確(E0001)或另開工作頁籤。";
            request.setAttribute("msg", Msg);
        } finally {
            dbc.closeCon();
        }
    }
    
    /**Purpose: 產生STU_CLASS的Excel下載範例
     * Author : Kimberly
     * Reivse : 20170626|Kimberly|在Servlet進行下載檔案的功能
     *          20180320|Kimberly|增加示範資料
     *          20180521|Kimberly|修改產生excel資料的寫法,原本是一個一個add到arrayList中,現在改成先將String字串們存在array中
     *          20210312|Kim     |新增頁籤：各學年課綱資料表；修改頁籤:[原]學制代碼表，調整為「學制及科別班群一覽表」
     */
    private void doPostExportClassExcel(HttpServletRequest request, 
                                        HttpServletResponse response, 
                                        String sbjYear,
                                        String sbjSem) throws ServletException, IOException 
    {
        ServletOutputStream         out = null;
        DBCon                dbc = null;
        Connection          conn = null;
        
        try
        {
            // 收集資料**
            request.setCharacterEncoding("UTF-8");      // Ajax need
            HttpSession session = request.getSession(); // 取得 session 物件
            
            // 設定資料**
            UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud
            String Sch_code = ud.getSch_code();         // 取得學校代碼
            int Sbj_year = ud.getSbj_year();            // 取得學年度
            String adg_code = ud.getAdg_code();         // 取得學校部別代碼
            
            dbc  = new DBCon(Sch_code);
            conn = dbc.getConnection();
            
            //設定下載EXCEL檔的檔名
            String fileName = sbjYear+"學年度第"+ sbjSem + "學期班級代碼表範例檔.xls";
            fileName = new String(fileName.getBytes(), "ISO8859-1");
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-Disposition","attachment;filename=\"" + fileName + "\"");
            
            //產生EXCEL的內容
            /***** 產生Sheet內容開始 *******************************/
            LinkedHashMap<String,ArrayList<ArrayList<String>>> sheetNameMap = new LinkedHashMap<String,ArrayList<ArrayList<String>>>();
//            ArrayList<SheetGenerater> dropdownDataList = new ArrayList<SheetGenerater>();
//            ArrayList<SheetGenerater> tooltipDataList = new ArrayList<SheetGenerater>();
            StuClass classdao = new StuClass();
            boolean isEmptySheet = false; //為了讓空白頁產生範例資料
            
            /**** 產生Sheet的區塊 ****/
            // 第一張Sheet: 產生SheetName與Sheet內容
            ArrayList<ArrayList<String>> sheetContentlist1 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList1 = new ArrayList<String>();
            String sheetName1 = "班級代碼表";
            String[] colsName= {"*學年","*學期","*班級代碼","*學制","*班群代碼","*教育部科別代碼","*年級","*班級全名(幾年幾班)"};
            rowList1= new ArrayList<String>(Arrays.asList(colsName));
            sheetContentlist1.add(rowList1);
            ArrayList<ArrayList<String>> classList = classdao.getClassDataAll(conn, sbjYear, sbjSem, adg_code);
            if(classList!=null && classList.size()>0) 
            {
                sheetContentlist1.addAll(classList);
            }
            else
            { 
                rowList1 = new ArrayList<String>();
                String[] colsExData= {"106","1","101","A","10","101","1","一年1班"};
                rowList1= new ArrayList<String>(Arrays.asList(colsExData));
                sheetContentlist1.add(rowList1);
                isEmptySheet = true;
                
            }
            sheetNameMap.put(sheetName1,sheetContentlist1);
            
            // 第二張Sheet
            ArrayList<ArrayList<String>> sheetContentlist2 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList2 = new ArrayList<String>();
            
            String[] colsExData= {"教育部科別代碼","教育部科別名稱","班群代碼","班群名稱","學年度"};
            rowList2= new ArrayList<String>(Arrays.asList(colsExData));
            sheetContentlist2.add(rowList2);
            
            if(classdao.getCprogramData(conn, Sbj_year, adg_code)!=null)
            	sheetContentlist2.addAll(classdao.getCprogramData(conn, Sbj_year, adg_code));
            
            sheetNameMap.put("參考(1)各學年課綱資料表 ",sheetContentlist2);
            
            // 第三張Sheet //1060920 Daphne
            ArrayList<ArrayList<String>> sheetContentlist3 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList3 = new ArrayList<String>();
            rowList3.add("教育部科別代碼");
            rowList3.add("教育部科別名稱");
            sheetContentlist3.add(rowList3);
            if(classdao.getDepartmentData(conn, adg_code)!=null)
                sheetContentlist3.addAll(classdao.getDepartmentData(conn, adg_code));
            sheetNameMap.put("參考(2)教育部科別代碼 ",sheetContentlist3);
            
            // 第四張Sheet //20210310 Kim
            ArrayList<ArrayList<String>> sheetContentlist4 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList4 = new ArrayList<String>();
            rowList4.add("學制代碼");
            rowList4.add("學制名稱");
            
            sheetContentlist4.add(rowList4);
            if(classdao.getMatricData(conn)!=null) {
                sheetContentlist4.addAll(classdao.getMatricData(conn));
            }
            
            String[] colsExData1= {"教育部科別代碼","教育部科別名稱","班群代碼","班群名稱"};
            rowList4= new ArrayList<String>(Arrays.asList(colsExData1));
            sheetContentlist4.add(rowList4);
            if(classdao.getDivisionData(conn, adg_code)!=null)
                sheetContentlist4.addAll(classdao.getDivisionData(conn, adg_code));
            sheetNameMap.put("參考(3)學制及科別班群一覽表",sheetContentlist4);

            
            /**** 產生下拉選單區塊 ****/
//            dropdownDataList = null;
            
            
            /**** 產生提示黃框區塊 ****/
//            tooltipDataList = null;

            
            /* 產生Excel檔 */
            HSSFWorkbook workbook = new HSSFWorkbook();  
            String SheetName = "";
            HSSFSheet sheet = null; 
            ArrayList<ArrayList<String>> sheetDataList = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList = new ArrayList<String>();
            
            //宣告POI物件
            HSSFRow row = null;
            HSSFCell cell = null;
            HSSFCellStyle styleheader;
            HSSFCellStyle stylecolumn;
            HSSFCellStyle styleExamplecolumn;
            
            /******* 設定標題單元格格式  *******/
            /* 設定文件格式 */
            //設定字型
            Font font = workbook.createFont();
            font.setColor(HSSFColor.BLACK.index);             //顏色
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            
            Font font2 = workbook.createFont();
            font2.setColor(HSSFColor.RED.index);             //顏色
            font2.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            //設定儲存格格式(例如:顏色)
            styleheader = workbook.createCellStyle();
            styleheader.setFont(font);
            styleheader.setFillForegroundColor(HSSFColor.PALE_BLUE.index); //設定顏色                                    
            styleheader.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleheader.setFillPattern((short) 1);
            //設定儲存格格線
            styleheader.setBorderBottom((short) 1);
            styleheader.setBorderTop((short) 1);
            styleheader.setBorderLeft((short) 1);
            styleheader.setBorderRight((short) 1);

            /* 設定內容文字格式 */
            stylecolumn = workbook.createCellStyle();
            stylecolumn.setBorderBottom((short) 1);
            stylecolumn.setBorderTop((short) 1);
            stylecolumn.setBorderLeft((short) 1);
            stylecolumn.setBorderRight((short) 1);
            
            /* 設定範例資料格式 */
            Font exampleFont = workbook.createFont();
            exampleFont.setColor(HSSFColor.GREY_50_PERCENT.index);
            styleExamplecolumn = workbook.createCellStyle();
            styleExamplecolumn.setFont(exampleFont);
            styleExamplecolumn.setFillForegroundColor(HSSFColor.WHITE.index); //設定顏色           
            styleExamplecolumn.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleExamplecolumn.setBorderBottom((short) 1);
            styleExamplecolumn.setBorderTop((short) 1);
            styleExamplecolumn.setBorderLeft((short) 1);
            styleExamplecolumn.setBorderRight((short) 1);
            
            
            //第一層for:有幾張sheet(Map)
            //第二層for:有幾列(ArrayList)
            //第三層for:有幾行
            int sheetNum = 0;
            for (Object key : sheetNameMap.keySet())
            {
                SheetName = key.toString();
                sheet = workbook.createSheet(SheetName);
               
               // System.out.println("SheetName==>" + SheetName);
                
               // System.out.println("sheetNum==>" + sheetNum);
                
                sheetDataList = sheetNameMap.get(key);
                for(int r = 0; r<sheetDataList.size();r++)
                {
                	row = sheet.createRow(r); 
                	rowList = sheetDataList.get(r); 
                	//System.out.println("rowList==>" + r);
                    for(int c = 0;c<rowList.size();c++)
                    {
                        cell = row.createCell(c); 
                        cell.setCellValue(rowList.get(c)); 
                       
                        if(r==0)
                        {
                            cell.setCellStyle(styleheader);                                                  
                            String str = rowList.get(c);
                            if (str.startsWith("*")) {
                                HSSFRichTextString richString = new HSSFRichTextString(str);
                                richString.applyFont(font2);
                                cell.setCellValue(richString);                            
                            }
                        }
                        else
                        {
                            if(sheetNum==0) {
                                if (isEmptySheet)
                                {
                                    cell.setCellStyle(styleExamplecolumn);
                                }
                                else
                                {
                                    cell.setCellStyle(stylecolumn);
                                }
                            }
                            else
                            {
                                cell.setCellStyle(stylecolumn);                               
                            }
                        }
                        // 2019-11-25增加同Sheet不同資料表欄位標頭顏色
						if (sheetNum == 1 || sheetNum == 3) {
							if (cell.getStringCellValue() == "教育部科別代碼" || cell.getStringCellValue() == "教育部科別名稱"
									|| cell.getStringCellValue() == "班群代碼" || cell.getStringCellValue() == "班群名稱"
									|| cell.getStringCellValue() == "學年度") {

								// System.out.println("cell.setCellValue=" + cell.getStringCellValue());
								cell.setCellStyle(styleheader);
							}
						}
                    }
                }                
                sheetNum++;
            }
            
            //設定標題列鎖定
            sheet.createFreezePane(0, 1, 0, 1);
            
            /***** 產生Sheet內容結束 *******************************/
            out = response.getOutputStream();
            workbook.write(out);
            workbook.close();
            logger.info("[school_roll] BTSchRoll doPostExportClassExcel success! " + ntpc.util.StringUtil.getLogInfo(request) + " sbjYear=" + sbjYear + " sbjSem=" + sbjSem);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
                if (dbc != null)
                    dbc.closeCon();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    
    /**Purpose: 產生STU_BASIS的Excel下載範例
     * Author : Josh
     * Reivse : 20170626|Kimberly|在Servlet進行下載檔案的功能
     *          20180320|Kimberly|新增頁籤：[原]科系所組代碼
     *          20180507|Kimberly|拿掉第3.4年及格分 & 第3.4年補考分
     *          20180515|Kimberly|新增學雜費減免身分&弱勢身分 + 新增特殊身分代碼表頁籤
     *          20180518|Kimberly|新增入學資格欄位
     *          20180521|Kimberly|修改產生excel資料的寫法,原本是一個一個add到arrayList中,現在改成先將String字串們存在array中
     *          20201104|Kim     |新增備註欄位
     *          20201124|Kim     |調整欄位順序
     */
    private void doPostExportBasisExcel(HttpServletRequest request,
                                        HttpServletResponse response, 
                                        String cmat_year) throws ServletException, IOException 
    {
        OutputStream out = null;
        DBCon                dbc = null;
        Connection          conn = null;
        
        try
        {
            // 收集資料**
            request.setCharacterEncoding("UTF-8");      // Ajax need
            HttpSession session = request.getSession(); // 取得 session 物件
            
            // 設定資料**
            UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud
            String Sch_code = ud.getSch_code();         // 取得學校代碼
            String adg_code = ud.getAdg_code();         // 取得學校部別代碼
            
            dbc  = new DBCon(Sch_code);
            conn = dbc.getConnection();
            
            //設定下載EXCEL檔的檔名
            String fileName = cmat_year + "學年度學生基本資料範例檔.xls";
            fileName = new String(fileName.getBytes(), "ISO8859-1");
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-Disposition","attachment;filename=\"" + fileName + "\"");
            
            //產生EXCEL的內容
            /***** 產生Sheet內容開始 *******************************/
            LinkedHashMap<String,ArrayList<ArrayList<String>>> sheetNameMap = new LinkedHashMap<String,ArrayList<ArrayList<String>>>();
//            ArrayList<SheetGenerater> dropdownDataList = new ArrayList<SheetGenerater>();
//            ArrayList<SheetGenerater> tooltipDataList = new ArrayList<SheetGenerater>();
            StuBasis basisDAO = new StuBasis();
            StuBasisDataList sbdl = new StuBasisDataList();
            boolean isEmptySheet = false; //為了讓空白頁產生範例資料
            
            /**** 產生Sheet的區塊 ****/
            // 第一張Sheet: 產生SheetName與Sheet內容
            ArrayList<ArrayList<String>> sheetContentlist1 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList1 = new ArrayList<String>();
            String sheetName2 = "學生基本資料表";
            String[] colsName = { "*學生學號","*中文姓名","*性別","*出生年月日","*護照種類","*身分證號",       //6
                                  "*課程適用年度","*學制","*班群代碼","*原入學班群代碼","*在學狀態代碼",        //11
                                  "*及格分類別代碼","*第一年及格分","*第二年及格分","*第三年及格分","*第一年補考分",//16
                                  "*第二年補考分","*第三年補考分","*監護人姓名","*入學資格學校代碼","*報部身份",   //21
                                  "*入學管道","*新生教育程度","*新生入學資格代碼","英文姓名","國籍","出生地",    //27
                                  "僑居地","是否中輟","戶籍地址縣市","戶籍地址鄉鎮","戶籍地址村里",
                                  "戶籍地址鄰","戶籍地址路街","戶籍地電話","通訊地址縣市","通訊地址鄉鎮",
                                  "通訊地址村里","通訊地址鄰","通訊地址路街","通訊地電話","電子信箱",
                                  "行動電話","血型","山地平地","原住民族別","畢業學校",
                                  "監護人關係","家長職業別","監護人地址","監護人電話(宅)","監護人電話(公)","監護人行動電話",
                                  "學雜費減免","弱勢身份","新生入學核准文字號","新生入學核准日期","母親身份證字號","父親身份證字號","備註" };

            rowList1= new ArrayList<String>(Arrays.asList(colsName));
            sheetContentlist1.add(rowList1);
            
            ArrayList<ArrayList<String>> basisList = basisDAO.getBasisDataAll(conn, cmat_year, adg_code);
            if(basisList!=null && basisList.size()>0) 
            {
                sheetContentlist1.addAll(basisList);
            }
            else
            { 
                rowList1 = new ArrayList<String>();
                String[] colsExData = { "1070101","示範者","1","1999-01-01","1","A123456789",
                                        "104","A","10","10","0",
                                        "0","60","60","60","40",
                                        "40","40","王爸爸","014362","0",
                                        "F","81","001","Shih,fan-che","台灣","台北市",
                                        "馬來西亞","","新北市","(可省略詳細資料填入地址路街此欄位)","(可省略詳細資料填入地址路街此欄位)",
                                        "1","XX鄉XX村XX路XX號","0912111111","新北市","(可省略詳細資料填入地址路街此欄位)",
                                        "(可省略詳細資料填入地址路街此欄位)","1","XX鄉XX村XX路XX號","0912111111","XXX@mail",
                                        "0912111111","A","1","01","測試國中",
                                        "父子","無業","XX市XX區XX路100號","02111111","02111111","0912111111",
                                        "0","0","※※※字第000001號","2018-01-01","A212345678","A112345678","" };
                                        //家長職業別
                rowList1= new ArrayList<String>(Arrays.asList(colsExData));
                
                sheetContentlist1.add(rowList1);
                isEmptySheet = true;
            }
            sheetNameMap.put(sheetName2,sheetContentlist1);
            
            
            // 第二張Sheet   Kimberly新增20180207  Daphne修改修改從DB撈出20180320
            ArrayList<ArrayList<String>> sheetContentlist2 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList2 = null;
            rowList2 = new ArrayList<String>();
            rowList2.add("代碼");
            rowList2.add("教育程度");
            sheetContentlist2.add(rowList2);
            sheetContentlist2.addAll(basisDAO.queryEduDegree(conn));
            sheetNameMap.put("參考(1)教育程度代碼 ",sheetContentlist2);
            
            
            //第三張Sheet  Josh新增20180312  Daphne修改從DB撈出20180320
            ArrayList<ArrayList<String>> sheetContentlist3 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList3 = null;
            rowList3 = new ArrayList<String>();
            rowList3.add("入學管道代碼");
            rowList3.add("入學管道名稱");
            sheetContentlist3.add(rowList3);           
            sheetContentlist3.addAll(basisDAO.queryAdmission(conn));
            sheetNameMap.put("參考(2)入學管道代碼 ",sheetContentlist3);
            
            
            // 第四張Sheet：參考(3)原班群代碼 
            ArrayList<ArrayList<String>> sheetContentlist4 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList4 = new ArrayList<String>();
            rowList4.add("班群代碼");
            rowList4.add("班群名稱");
            sheetContentlist4.add(rowList4);
            if(sbdl.getDivisionData(conn)!=null)
                sheetContentlist4.addAll(sbdl.getDivisionData(conn));
            sheetNameMap.put("參考(3)原班群代碼 ",sheetContentlist4);
            
            
            // 第五張Sheet：參考(4)特殊身份代碼
            ArrayList<ArrayList<String>> sheetContentlist5 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList5 = new ArrayList<String>();
            rowList5.add("特殊身份名稱");
            rowList5.add("代碼"); 
            sheetContentlist5.add(rowList5);
            if(sbdl.getDivisionData(conn)!=null)
                sheetContentlist5.addAll(sbdl.getStuIdCode(conn));
            sheetNameMap.put("參考(4)特殊身份代碼",sheetContentlist5);
            
            
         // 第六張Sheet：參考(5)入學資格學校代碼
            ArrayList<ArrayList<String>> sheetContentlist6 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList6 = new ArrayList<String>();
            rowList6.add("學校代碼");
            rowList6.add("學校名稱"); 
            sheetContentlist6.add(rowList6);
            sheetContentlist6.addAll(sbdl.getEnterSch(conn));
            sheetNameMap.put("參考(5)入學資格學校代碼",sheetContentlist6);
            
         // 第七張Sheet：參考(6)原住民族別代碼
            ArrayList<ArrayList<String>> sheetContentlist7 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList7 = new ArrayList<String>();
            rowList7.add("族別代碼");
            rowList7.add("族別名稱"); 
            sheetContentlist7.add(rowList7);
            sheetContentlist7.addAll(sbdl.getTwNative(conn));
            sheetNameMap.put("參考(6)原住民族別代碼",sheetContentlist7);
            
         // 第八張Sheet：參考(7)及格分類代碼
            ArrayList<ArrayList<String>> sheetContentlist8 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList8 = new ArrayList<String>();
            rowList8.add("及格分類代碼");
            rowList8.add("身分類別名稱");
            rowList8.add("第一年及格分");
            rowList8.add("第二年及格分");
            rowList8.add("第三年及格分");
            rowList8.add("第一年補考及格分");
            rowList8.add("第二年補考及格分");
            rowList8.add("第三年補考及格分");
            sheetContentlist8.add(rowList8);
            sheetContentlist8.addAll(sbdl.getSpeclib(conn));
            sheetNameMap.put("參考(7)及格分類代碼",sheetContentlist8);
            
         // 第九張Sheet：參考(8)新生入學資格代碼
            ArrayList<ArrayList<String>> sheetContentlist9 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList9 = new ArrayList<String>();
            rowList9.add("入學資格代碼");
            rowList9.add("入學資格名稱");
            sheetContentlist9.add(rowList9);
            sheetContentlist9.addAll(sbdl.getAdmissionQualification(conn));
            sheetNameMap.put("參考(8)新生入學資格代碼",sheetContentlist9);
            
         // 第十張Sheet：參考(9)縣市鄉鎮村里一覽表
            ArrayList<ArrayList<String>> sheetContentlist10 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList10 = new ArrayList<String>();
            rowList10.add("地址縣市");
            rowList10.add("地址鄉鎮");
            rowList10.add("地址村里");
            sheetContentlist10.add(rowList10);
            sheetContentlist10.addAll(sbdl.getCityTownVillage(conn));
            sheetNameMap.put("參考(9)縣市鄉鎮村里一覽表",sheetContentlist10);
            
         // 第十一張Sheet：參考(10)家長職業別一覽表
            ArrayList<ArrayList<String>> sheetContentlist11 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList11 = new ArrayList<String>();
            rowList11.add("家長職業別");
            sheetContentlist11.add(rowList11);

            String occupation[] = {"軍公教人員","商業","工業","農林漁牧業","醫療業","服務業","家管","自由業","其他(自行輸入)"};
            for (int i = 0; i < occupation.length; i++) {
                ArrayList<String> newRow = new ArrayList<String>();
                newRow.add(occupation[i]);
                sheetContentlist11.add(newRow);
            }

            sheetNameMap.put("參考(10)家長職業別一覽表",sheetContentlist11);
            
            /**** 產生下拉選單區塊 ****/
//            dropdownDataList = null;
            
            /**** 產生提示黃框區塊 ****/
//            tooltipDataList = null;
            
            
            /* 產生Excel檔 */
            HSSFWorkbook workbook = new HSSFWorkbook();  
            String SheetName = "";
            HSSFSheet sheet = null; 
            ArrayList<ArrayList<String>> sheetDataList = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList = new ArrayList<String>();
            
            //宣告POI物件
            HSSFRow row = null;
            HSSFCell cell = null;
            HSSFCellStyle styleheader;
            HSSFCellStyle stylecolumn;
            HSSFCellStyle styleExamplecolumn;
            
            /******* 設定標題單元格格式  *******/
            /* 設定文件格式 */
            //設定字型
            Font font = workbook.createFont();
            font.setColor(HSSFColor.BLACK.index);             //顏色
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            
            Font font2 = workbook.createFont();
            font2.setColor(HSSFColor.RED.index);             //顏色
            font2.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            //設定儲存格格式(例如:顏色)
            styleheader = workbook.createCellStyle();
            
            styleheader.setFont(font);
            styleheader.setFillForegroundColor(HSSFColor.PALE_BLUE.index); //設定顏色                                    
            styleheader.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleheader.setFillPattern((short) 1);
            //設定儲存格格線
            styleheader.setBorderBottom((short) 1);
            styleheader.setBorderTop((short) 1);
            styleheader.setBorderLeft((short) 1);
            styleheader.setBorderRight((short) 1);
    
            /* 設定內容文字格式 */
            stylecolumn = workbook.createCellStyle();
            stylecolumn.setBorderBottom((short) 1);
            stylecolumn.setBorderTop((short) 1);
            stylecolumn.setBorderLeft((short) 1);
            stylecolumn.setBorderRight((short) 1);
            
            /* 設定範例資料格式 */
            Font exampleFont = workbook.createFont();
            exampleFont.setColor(HSSFColor.GREY_50_PERCENT.index);
            styleExamplecolumn = workbook.createCellStyle();
            styleExamplecolumn.setFont(exampleFont);
            styleExamplecolumn.setFillForegroundColor(HSSFColor.WHITE.index); //設定顏色           
            styleExamplecolumn.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleExamplecolumn.setBorderBottom((short) 1);
            styleExamplecolumn.setBorderTop((short) 1);
            styleExamplecolumn.setBorderLeft((short) 1);
            styleExamplecolumn.setBorderRight((short) 1);
            
            
            //第一層for:有幾張sheet(Map)
            //第二層for:有幾列(ArrayList)
            //第三層for:有幾行
            int sheetNum = 0;
            for (Object key : sheetNameMap.keySet())
            {
                SheetName = key.toString();
                sheet = workbook.createSheet(SheetName);
                
//                System.out.println("SheetName==>" + SheetName);
//                System.out.println("sheet==>" + sheet);
                
                sheetDataList = sheetNameMap.get(key);
                for(int r = 0; r<sheetDataList.size();r++)
                {
                    row = sheet.createRow(r); 
                    rowList = sheetDataList.get(r);
                    for(int c = 0;c<rowList.size();c++)
                    {
                        cell = row.createCell(c); 
                        cell.setCellValue(rowList.get(c));  
                        if(r==0)
                        {
                            cell.setCellStyle(styleheader);
                            
                            String str = rowList.get(c);
                            if (str.startsWith("*")) {
                                HSSFRichTextString richString = new HSSFRichTextString(str);
                                richString.applyFont(font2);
                                cell.setCellValue(richString);
                            }
                        }
                        else
                        {
                            if(sheetNum==0) {
                                if (isEmptySheet)
                                {
                                    cell.setCellStyle(styleExamplecolumn);
                                }
                                else
                                {
                                    cell.setCellStyle(stylecolumn);
                                }
                            }
                            else
                            {
                                cell.setCellStyle(stylecolumn);                               
                            }
                        }
                    }
                }
                sheetNum++;
            }
            
            /**** 產生連動式下拉選單區塊 ****/
            // 所有縣市名稱
            ArrayList<String> CityName = sbdl.getAddrCityCname(conn);       //縣市代碼名稱
            String[] provNameList = CityName.toArray(new String[0]);   
            
            // 整理數據，放入一個map中，mapkey存放縣市，value存放該地點下的子區域
            LinkedHashMap<String, List<String>> siteMap = new LinkedHashMap<String, List<String>>();
            ArrayList<ArrayList<String>> CityCodeMap       = sbdl.getAddrCity(conn);       //縣市代碼
            for (int x = 0; x < CityCodeMap.size(); x++) {
            	ArrayList<ArrayList<String>> TownMap        = sbdl.getAddrTown(conn, CityCodeMap.get(x).get(0)); //鄉鎮代碼
            	ArrayList<String> townCname = new ArrayList<String>();
            	for (int m = 0; m < TownMap.size(); m++) {
            		townCname.add(TownMap.get(m).get(1));
            	}
            	siteMap.put(CityCodeMap.get(x).get(1), townCname); //縣市名稱
//            	ArrayList<String> townNum = new ArrayList<String>();
            	for (int y = 0; y < TownMap.size(); y++) {
            		ArrayList<String> VillageMap        = sbdl.getAddrVillage(conn, TownMap.get(y).get(0)); //村里代碼
            		siteMap.put(TownMap.get(y).get(1), VillageMap);
            	}
            }
            
            // 創建一個專門用來存放地區資料的隱藏sheet頁
            HSSFSheet hideSheet = workbook.createSheet("site_sheet");
            // 這一行作用是將此sheet隱藏，功能未完成時註釋此行,可以查看隱藏sheet中信息是否正確
            workbook.setSheetHidden(workbook.getSheetIndex(hideSheet), true);
            
            int rowId = 0;
            // 設置第一行，存縣市的資料
            HSSFRow provinceRow = hideSheet.createRow(rowId++);  
            provinceRow.createCell(0).setCellValue("縣市列表");  
            for(int i = 0; i < provNameList.length; i ++){  
                HSSFCell provinceCell = provinceRow.createCell(i + 1);  
                provinceCell.setCellValue(provNameList[i]);  
            }
            // 將具體的數據寫入到每一行中，行開頭爲父級區域，後面是子區域。
            Iterator<String> keyIterator = siteMap.keySet().iterator();
            while(keyIterator.hasNext()) {
            	String key = keyIterator.next();
            	List<String> son = siteMap.get(key);
            	
            	row = hideSheet.createRow(rowId++);  
            	row.createCell(0).setCellValue(key);
            	for(int j = 0; j < son.size(); j ++){  
            		cell = row.createCell(j + 1);  
            		cell.setCellValue(son.get(j));  
            	}
            	
            	// 添加名稱管理器  
            	String range = getRange(1, rowId, son.size());
            	HSSFName name = workbook.createName();
            	//key不可重複,將父區域名作爲key
            	name.setNameName(key);
            	String formula = "site_sheet!" + range;  
            	name.setRefersToFormula(formula);
            }
            
            HSSFSheet sheetPro = workbook.getSheet(sheetName2); //指定現有工作表
            
            HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper((HSSFSheet)sheetPro);  
            // 省規則  
            DVConstraint provConstraint = DVConstraint.createExplicitListConstraint(provNameList); 
            // 四個參數分別是：起始行、終止行、起始列、終止列
            CellRangeAddressList provRangeAddressList = new CellRangeAddressList(1, 2399, 29, 29);
            DataValidation provinceDataValidation = dvHelper.createValidation(provConstraint, provRangeAddressList);
            CellRangeAddressList provRangeComAddressList = new CellRangeAddressList(1, 2399, 35, 35);
            DataValidation provinceComDataValidation = dvHelper.createValidation(provConstraint, provRangeComAddressList);  
            // 驗證
            provinceDataValidation.createErrorBox("error", "請選擇正確的縣市");  
            provinceDataValidation.setShowErrorBox(true);          // 設置顯示錯誤框
//            provinceDataValidation.setSuppressDropDownArrow(true); // 設置抑制下拉箭頭
//            provinceDataValidation.setSuppressDropDownArrow(false);
            sheetPro.addValidationData(provinceDataValidation);
            
            provinceComDataValidation.createErrorBox("error", "請選擇正確的縣市");  
            provinceComDataValidation.setShowErrorBox(true);       // 設置顯示錯誤框
            sheetPro.addValidationData(provinceComDataValidation);
            
            //對前2400行設置有效性
            for(int i = 2; i < 2401; i++){
            	// 戶籍地址
                setDataValidation("AD", sheetPro, i, 31);
                setDataValidation("AE", sheetPro, i, 32);
                // 通訊地址
                setDataValidation("AJ", sheetPro, i, 37);
                setDataValidation("AK", sheetPro, i, 38);
            }
                        
            /**** 產生下拉選單區塊 ****/
//            String[] city = {"台北市", "新北市", "宜蘭市", "台中市"};
//            String[] textlist = {"大安區", "大同區", "信義區"};
//            sheet = ParseXLSUtil.setHSSFValidation(sheetPro, city, 1, 11, 0, 0);
//            sheet = ParseXLSUtil.setHSSFValidation(sheetPro, textlist, 1, 11, 1, 1);
            
            XSSFSheet sheet2 = workbook.getSheet(sheetName2);

            // 在第 AW 行下方創建下拉式選單
            int rowIndex = sheet2.getLastRowNum() + 1; // 最後一行的下一行
            int columnIndex = 49; // 第 AW 列

            // 設置下拉式選單的選項值
            String[] options = {"軍公教人員","商業","工業","農林漁牧業","醫療業","服務業","家管","自由業","其他(自行輸入)"};
            DataValidationHelper validationHelper = new XSSFDataValidationHelper(sheet2);
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
            CellRangeAddressList addressList = new CellRangeAddressList(rowIndex, rowIndex, columnIndex, columnIndex);
            DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
            
            //設定標題列鎖定
            sheet.createFreezePane(0, 1, 0, 1);
            
            /***** 產生Sheet內容結束 *******************************/
            out = response.getOutputStream();
            workbook.write(out);
            workbook.close();
            logger.info("[school_roll] BTSchRoll doPostExportBasisExcel success! " + ntpc.util.StringUtil.getLogInfo(request) + " cmat_year=" + cmat_year);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
                if (dbc != null)
                    dbc.closeCon();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    
    /**Purpose: 產生STU_REGISTER的Excel下載範例
     * Author : Daphne
     * Reivse : 20170626|Kimberly|在Servlet進行下載檔案的功能
     *          20180320|Kimberly|增加示範資料
     *          20180521|Kimberly|修改產生excel資料的寫法,原本是一個一個add到arrayList中,現在改成先將String字串們存在array中
     */
    private void doPostExportRegisterExcel(HttpServletRequest request, 
                                           HttpServletResponse response, 
                                           String sbjYear,
                                           String sbjSem) throws ServletException, IOException 
    {
        OutputStream out = null;
        DBCon                dbc = null;
        Connection          conn = null;
        
        try
        {
            // 收集資料**
            request.setCharacterEncoding("UTF-8");      // Ajax need
            HttpSession session = request.getSession(); // 取得 session 物件
            
            // 設定資料**
            UserData ud = (UserData)session.getAttribute("ud");  // 取得 session 中的ud
            String Sch_code = ud.getSch_code();         // 取得學校代碼
            String adg_code = ud.getAdg_code();         // 取得學校部別代碼
            
            dbc  = new DBCon(Sch_code);
            conn = dbc.getConnection();
        
            //設定下載EXCEL檔的檔名
            String fileName = sbjYear+"學年度第"+ sbjSem + "學期註冊編班資料範例檔.xls";
            fileName = new String(fileName.getBytes(), "ISO8859-1");
            response.setCharacterEncoding("UTF-8");
            response.setHeader("Content-Disposition","attachment;filename=\"" + fileName + "\"");
            
            //產生EXCEL的內容
            /***** 產生Sheet內容開始 *******************************/
            LinkedHashMap<String,ArrayList<ArrayList<String>>> sheetNameMap = new LinkedHashMap<String,ArrayList<ArrayList<String>>>();
//            ArrayList<SheetGenerater> dropdownDataList = new ArrayList<SheetGenerater>();
//            ArrayList<SheetGenerater> tooltipDataList = new ArrayList<SheetGenerater>();
            StuRegister registerDAO = new StuRegister();
            boolean isEmptySheet = false;
            
            // 第一張Sheet: 產生SheetName與Sheet內容
            ArrayList<ArrayList<String>> sheetContentlist1 = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList1 = new ArrayList<String>();
            String sheetName1 = "註冊編班資料表";
            String[] colsName = { "*學年度","*學期","中文姓名","*身分證字號","*學生學號","*班級代碼",
                                  "*班級座號","*學期在學狀態代碼","*註冊記錄","休學記錄序號","復學生註記",
                                  "重讀生註記","轉學轉科註記","*班群代碼" };
            rowList1= new ArrayList<String>(Arrays.asList(colsName));
            sheetContentlist1.add(rowList1);
            
            ArrayList<ArrayList<String>> registerList = registerDAO.getRegisterDataAll(conn, sbjYear, sbjSem, adg_code);
            if(registerList!=null && registerList.size()>0) {                
                sheetContentlist1.addAll(registerList);
                
            } else { 
                rowList1 = new ArrayList<String>();
                String[] colsExData = { "104","1","非必填","A123456789","1070101","A01",
                                        "01","0","2","0","0",
                                        "0","0","11" };
                rowList1= new ArrayList<String>(Arrays.asList(colsExData));
                sheetContentlist1.add(rowList1);
                isEmptySheet = true;       
                
            }
            sheetNameMap.put(sheetName1,sheetContentlist1);
            
            /**** 產生下拉選單區塊 ****/
//            dropdownDataList = null;
            
            /**** 產生提示黃框區塊 ****/
//            tooltipDataList = null;
            
            /* 產生Excel檔 */
            HSSFWorkbook workbook = new HSSFWorkbook();  
            String SheetName = "";
            HSSFSheet sheet = null; 
            ArrayList<ArrayList<String>> sheetDataList = new ArrayList<ArrayList<String>>();
            ArrayList<String> rowList = new ArrayList<String>();
            
            //宣告POI物件
            HSSFRow row = null;
            HSSFCell cell = null;
            HSSFCellStyle styleheader;
            HSSFCellStyle stylecolumn;
            HSSFCellStyle styleExamplecolumn;
            
            /* 設定文件格式 */
            //設定字型
            Font font = workbook.createFont();
            font.setColor(HSSFColor.BLACK.index);             //顏色
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            
            Font font2 = workbook.createFont();
            font2.setColor(HSSFColor.RED.index);             //顏色
            font2.setBoldweight(Font.BOLDWEIGHT_BOLD);         //粗體
            
            //設定儲存格格式(例如:顏色)
            styleheader = workbook.createCellStyle();
            styleheader.setFont(font);
            styleheader.setFillForegroundColor(HSSFColor.PALE_BLUE.index); //設定顏色                                    
            styleheader.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleheader.setFillPattern((short) 1);
            
            //設定儲存格格線
            styleheader.setBorderBottom((short) 1);
            styleheader.setBorderTop((short) 1);
            styleheader.setBorderLeft((short) 1);
            styleheader.setBorderRight((short) 1);
            
            /* 設定內容文字格式 */
            stylecolumn = workbook.createCellStyle();
            stylecolumn.setBorderBottom((short) 1);
            stylecolumn.setBorderTop((short) 1);
            stylecolumn.setBorderLeft((short) 1);
            stylecolumn.setBorderRight((short) 1);
                        
            /* 設定範例資料格式 */
            Font exampleFont = workbook.createFont();
            exampleFont.setColor(HSSFColor.GREY_50_PERCENT.index);
            styleExamplecolumn = workbook.createCellStyle();
            styleExamplecolumn.setFont(exampleFont);
            styleExamplecolumn.setFillForegroundColor(HSSFColor.WHITE.index); //設定顏色           
            styleExamplecolumn.setAlignment(HSSFCellStyle.ALIGN_CENTER);          //水平置中 
            styleExamplecolumn.setBorderBottom((short) 1);
            styleExamplecolumn.setBorderTop((short) 1);
            styleExamplecolumn.setBorderLeft((short) 1);
            styleExamplecolumn.setBorderRight((short) 1);            
            
            //第一層for:有幾張sheet(Map)
            //第二層for:有幾列(ArrayList)
            //第三層for:有幾行
            int sheetNum = 0;
            for (Object key : sheetNameMap.keySet())
            {
                SheetName = key.toString();
                sheet = workbook.createSheet(SheetName);
                
                sheetDataList = sheetNameMap.get(key);
                for(int r = 0; r<sheetDataList.size();r++)
                {
                    row = sheet.createRow(r); 
                    rowList = sheetDataList.get(r);
                    for(int c = 0;c<rowList.size();c++)
                    {
                        cell = row.createCell(c); 
                        cell.setCellValue(rowList.get(c));  
                        if(r==0)
                        {
                            cell.setCellStyle(styleheader);
                            
                            String str = rowList.get(c);
                            if (str.startsWith("*")) {
                                HSSFRichTextString richString = new HSSFRichTextString(str);
                                richString.applyFont(font2);
                                cell.setCellValue(richString);
                            }
                        }
                        else
                        {
                            if(sheetNum==0) {
                                if (isEmptySheet)
                                {
                                    cell.setCellStyle(styleExamplecolumn);
                                }
                                else
                                {
                                    cell.setCellStyle(stylecolumn);
                                }
                            }
                            else
                            {
                                cell.setCellStyle(stylecolumn);                               
                            }
                        }
                    }
                }
                sheetNum++;
            }
            //設定標題列鎖定
            sheet.createFreezePane(0, 1, 0, 1);
            /***** 產生Sheet內容結束 *******************************/
            
            out = response.getOutputStream();
            workbook.write(out);
            workbook.close();
            logger.info("[school_roll] BTSchRoll doPostExportRegisterExcel success! " + ntpc.util.StringUtil.getLogInfo(request) + " sbjYear=" + sbjYear + " sbjSem=" + sbjSem);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.flush();
                    out.close();
                }
                if (dbc != null)
                    dbc.closeCon();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    
    /**
     * 設置有效性
     * @param offset 主影響單元格所在列，即此單元格由哪個單元格影響聯動
     * @param sheet
     * @param rowNum 行數
     * @param colNum 列數
     */
    public static void setDataValidation(String offset, HSSFSheet sheet, int rowNum, int colNum) {
        HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper(sheet);
        HSSFDataValidation data_validation_list;
            data_validation_list = getDataValidationByFormula(
            		"INDIRECT($" + offset + (1) + ")", dvHelper, rowNum, colNum);
        sheet.addValidationData(data_validation_list);
    }

    /**
     * 加載下拉列表內容
     * @param formulaString
     * @param naturalRowIndex
     * @param naturalColumnIndex
     * @param dvHelper
     * @return
     */
    private static  HSSFDataValidation getDataValidationByFormula(
            String formulaString, HSSFDataValidationHelper dvHelper, int naturalRowIndex, int naturalColumnIndex) {
//        System.out.println("formulaString = " + formulaString + "; naturalRowIndex =" + naturalRowIndex + "; naturalColumnIndex =" + naturalColumnIndex + "; dvHelper = " + dvHelper);
    	// 加載下拉列表內容
        // 舉例：若formulaString = "INDIRECT($A$2)" 表示規則數據會從名稱管理器中獲取key與單元格 A2 值相同的數據，
        // 如果A2是江蘇省，那麼此處就是江蘇省下的市信息。 
    	DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint(formulaString);
        // 設置數據有效性加載在哪個單元格上。
        // 四個參數分別是：起始行、終止行、起始列、終止列
        int firstRow = naturalRowIndex - 1;
        int lastRow = naturalRowIndex - 1;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = naturalColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow,
                lastRow, firstCol, lastCol);
        // 數據有效性對象
        // 綁定
        HSSFDataValidation data_validation_list = (HSSFDataValidation) dvHelper.createValidation(dvConstraint, regions);
        data_validation_list.setEmptyCellAllowed(false);         // 設置允許空單元格
        //System.out.println("判斷為 =>" + (data_validation_list instanceof HSSFDataValidation));
        if (data_validation_list instanceof HSSFDataValidation) {
        	data_validation_list.setSuppressDropDownArrow(false);
        } else {
        	data_validation_list.setSuppressDropDownArrow(true); // 設置抑制下拉箭頭
        	data_validation_list.setShowErrorBox(true);          // 設置顯示錯誤框
        }
        
        // 設置輸入信息提示信息
        data_validation_list.createPromptBox("下拉選擇提示", "請使用下拉方式選擇合適的值！");
        // 設置輸入錯誤提示信息
        data_validation_list.createErrorBox("選擇錯誤提示", "你輸入的值未在備選列表中，請下拉選擇合適的值！");
        return data_validation_list;
    }
    
    /**
     *  計算formula
     * @param offset 偏移量，如果給0，表示從A列開始，1，就是從B列 
     * @param rowId 第幾行 
     * @param colCount 一共多少列 
     * @return 如果給入參 1,1,10. 表示從B1-K1。最終返回 $B$1:$K$1 
     *  
     */  
    public String getRange(int offset, int rowId, int colCount) {
        char start = (char)('A' + offset);  
        if (colCount <= 25) {  
            char end = (char)(start + colCount - 1);  
            return "$" + start + "$" + rowId + ":$" + end + "$" + rowId;  
        } else {  
            char endPrefix = 'A';  
            char endSuffix = 'A';  
            if ((colCount - 25) / 26 == 0 || colCount == 51) {// 26-51之間，包括邊界（僅兩次字母表計算）  
                if ((colCount - 25) % 26 == 0) {// 邊界值  
                    endSuffix = (char)('A' + 25);  
                } else {  
                    endSuffix = (char)('A' + (colCount - 25) % 26 - 1);  
                }  
            } else {// 51以上  
                if ((colCount - 25) % 26 == 0) {  
                    endSuffix = (char)('A' + 25);  
                    endPrefix = (char)(endPrefix + (colCount - 25) / 26 - 1);  
                } else {
                    endSuffix = (char)('A' + (colCount - 25) % 26 - 1);  
                    endPrefix = (char)(endPrefix + (colCount - 25) / 26);  
                }  
            }
            return "$" + start + "$" + rowId + ":$" + endPrefix + endSuffix + "$" + rowId;
            
        }  
    }  
}
