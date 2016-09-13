package com.acmr.mq.consumer.queue;

import org.apache.log4j.Logger;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.service.CellService;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.SheetService;
import com.acmr.mq.AffectCell;
import com.acmr.mq.Model;

public class WorkerThread2 implements Runnable{
	private static Logger logger = Logger.getLogger(QueueReceiver.class);
	private int step;  
	private MemcachedClient memcachedClient;
	private String key;
	private HandleExcelService handleExcelService;
	private CellService cellService;
	private PasteService pasteService;
	private SheetService sheetService;
	private Model model;
	
	
    
    public WorkerThread2(int step,MemcachedClient memcachedClient,String key,HandleExcelService handleExcelService,
    		CellService cellService,PasteService pasteService,SheetService sheetService,Model model){  
        this.step=step;
        this.memcachedClient = memcachedClient;
        this.key = key;
        this.handleExcelService = handleExcelService;
        this.cellService = cellService;
        this.pasteService = pasteService;
        this.sheetService = sheetService;
        this.model = model;
    }  
   
    @Override  
    public void run() {  
    	while(true){
    		int memStep = (Integer) memcachedClient.get(key);
    		if(memStep + 1 == step){
    			System.out.println(step + "开始执行");
    			logger.info("**********begin excelId : "+model.getExcelId() + " === step : " + step + "== reqPath : "+ model.getReqPath());
    			handleMessage(model);
    			return;
    		}else{
    			processCommand(10);
    			continue;
    		}
    	}
    }  
   
    private void processCommand(int n) {  
        try {  
            Thread.sleep(n);  
        } catch (InterruptedException e) {  
            e.printStackTrace();  
        }  
    }  
   
    @Override  
    public String toString(){  
        return this.step+"";  
    }
    private void handleMessage(Model model) {
		int reqPath = model.getReqPath();
		String excelId = model.getExcelId();
		int step = model.getStep();
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		Cell cell = null;
		//AffectCell affectCell = new AffectCell();
		switch (reqPath) {
		case OperatorConstant.textData:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getStartX());
			handleExcelService.data(cell, excelBook);
			break;
		case OperatorConstant.fontsize:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.font_size, cell,excelBook);
			break;
		case OperatorConstant.fontfamily:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.font_family, cell,excelBook);
			break;
		case OperatorConstant.fontweight:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.font_weight, cell,excelBook);
			break;
		case OperatorConstant.fontitalic:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.font_italic, cell,excelBook);
			break;
		case OperatorConstant.fontcolor:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.font_color, cell,excelBook);
			break;
		case OperatorConstant.wordWrap:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.word_wrap, cell,excelBook);
			break;

		case OperatorConstant.fillbgcolor:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook);
			break;
		case OperatorConstant.textDataformat:
			CellFormate cellFormate = (CellFormate) model.getObject();
//			affectCell.setStartRowAlias(cellFormate.getCoordinate().getStartRowAlais());
//			affectCell.setEndRowAlias(cellFormate.getCoordinate().getEndRowAlais());
//			affectCell.setStartColAlias(cellFormate.getCoordinate().getStartColAlais());
//			affectCell.setEndColAlias(cellFormate.getCoordinate().getEndColAlais());
			handleExcelService.setCellFormate(cellFormate, excelBook);
			break;

		case OperatorConstant.commentset:
			Comment comment = (Comment) model.getObject();
//			affectCell.setStartRowAlias(comment.getCoordinate().getStartRowAlais());
//			affectCell.setEndRowAlias(comment.getCoordinate().getEndRowAlais());
//			affectCell.setStartColAlias(comment.getCoordinate().getStartColAlais());
//			affectCell.setEndColAlias(comment.getCoordinate().getEndColAlais());
			handleExcelService.setComment(excelBook, comment);
			break;
		case OperatorConstant.merge:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			cellService.mergeCell(excelBook.getSheets().get(0), cell);
			break;
		case OperatorConstant.mergedelete:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			cellService.splitCell(excelBook.getSheets().get(0), cell);
			break;
		case OperatorConstant.frame:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.frame, cell,excelBook);
			break;
		case OperatorConstant.alignlevel:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.align_level, cell,excelBook);
			break;
		case OperatorConstant.alignvertical:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
			handleExcelService.updateCells(CellUpdateType.align_vertical, cell,excelBook);
			break;
		case OperatorConstant.rowsinsert:
			RowOperate rowOperate = (RowOperate) model.getObject();
			//affectCell.setStartRowAlias(rowOperate.getRowAlias());
//			affectCell.setEndRowAlias(rowOperate.getRowAlias());
//			affectCell.setStartColAlias("MIN");
//			affectCell.setEndColAlias("MAX");
			//affectCell.setType("rows_insert");
			cellService.addRow(excelBook.getSheets().get(0), rowOperate);
			break;
		case OperatorConstant.rowsdelete:
			RowOperate rowOperate2 = (RowOperate) model.getObject();
			//affectCell.setStartRowAlias(rowOperate2.getRowAlias());
//			affectCell.setEndRowAlias(rowOperate2.getRowAlias());
//			affectCell.setStartColAlias("MIN");
//			affectCell.setEndColAlias("MAX");
			//affectCell.setType("rows_delete");
			cellService.deleteRow(excelBook.getSheets().get(0), rowOperate2);
			break;
		case OperatorConstant.colsinsert:
			ColOperate colOperate = (ColOperate) model.getObject();
//			affectCell.setStartRowAlias("MIN");
//			affectCell.setEndRowAlias("MAX");
			//affectCell.setStartColAlias(colOperate.getColAlias());
			//affectCell.setEndColAlias(colOperate.getColAlias());
			//affectCell.setType("cols_insert");
			cellService.addCol(excelBook.getSheets().get(0), colOperate);
			break;
		case OperatorConstant.colsdelete:
			ColOperate colOperate2 = (ColOperate) model.getObject();
//			affectCell.setStartRowAlias("MIN");
//			affectCell.setEndRowAlias("MAX");
		//	affectCell.setStartColAlias(colOperate2.getColAlias());
			//affectCell.setEndColAlias(colOperate2.getColAlias());
			//affectCell.setType("cols_delete");
			cellService.deleteCol(excelBook.getSheets().get(0), colOperate2);
			break;
		case OperatorConstant.paste:
			Paste paste = (Paste) model.getObject();
//			affectCell.setStartRowAlias(paste.getStartRowAlias());
//			affectCell.setStartColAlias(paste.getStartColAlias());
			pasteService.data(paste, excelBook);
			break;
		case OperatorConstant.copy:
			Copy copy = (Copy) model.getObject();
//			affectCell.setStartRowAlias(copy.getTarget().getRowAlias());
//			affectCell.setStartColAlias(copy.getTarget().getColAlias());
			pasteService.copy(copy, excelBook);
			break;
		case OperatorConstant.cut:
			Copy copy2 = (Copy) model.getObject();
//			affectCell.setStartRowAlias(copy2.getTarget().getRowAlias());
//			affectCell.setStartColAlias(copy2.getTarget().getColAlias());
			pasteService.cut(copy2, excelBook);
			break;
		case OperatorConstant.frozen:
			Frozen frozen = (Frozen) model.getObject();
//			affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(frozen.getFrozenX());
			sheetService.frozen(excelBook.getSheets().get(0), frozen);
			break;
		case OperatorConstant.unFrozen:
//			affectCell.setStartRowAlias(frozen2.getFrozenY());
//			affectCell.setStartColAlias(frozen2.getFrozenX());
			excelBook.getSheets().get(0).setFreeze(null);
			break;
		case OperatorConstant.colswidth:
			ColWidth colWidth = (ColWidth) model.getObject();
			//affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(colWidth.getColAlias());
//			affectCell.setType("cols_width");
			cellService.controlColWidth(excelBook.getSheets().get(0), colWidth);
			break;
		case OperatorConstant.colshide:
			ColOperate colHide = (ColOperate) model.getObject();
			//affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(colWidth.getColAlias());
//			affectCell.setType("cols_width");
			cellService.colHide(excelBook.getSheets().get(0), colHide);
			break;	
		case OperatorConstant.colhideCancel:
			ColOperate colhideCancel = (ColOperate) model.getObject();
			//affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(colWidth.getColAlias());
//			affectCell.setType("cols_width");
			sheetService.cancelColHide(excelBook.getSheets().get(0), colhideCancel);
			break;	
		case OperatorConstant.rowsheight:
			RowHeight rowHeight = (RowHeight) model.getObject();
//			affectCell.setStartRowAlias(rowHeight.getRowAlias());
//			affectCell.setType("rows_height");
			cellService.controlRowHeight(excelBook.getSheets().get(0), rowHeight);
			break;
		case OperatorConstant.addRowLine:
			RowLine rowLine = (RowLine) model.getObject();
		//	affectCell.setStartRowAlias(rowLine.getRowNum());
			//affectCell.setStartColAlias(colWidth.getColAlias());
			String rowNum = rowLine.getRowNum();
			int rn = Integer.valueOf(rowNum);
			sheetService.addRowLine(excelBook.getSheets().get(0),rn);
			break;
		case OperatorConstant.colorset:
			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getStartX());
			handleExcelService.colorSet(cell, excelBook);	
//		case 29:
//			//Model mod =  (Model)model.getObject();
//			break;
		default:
			break;
		}
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME, excelBook);
		//memcachedClient.set(excelId+"_"+step, Constant.MEMCACHED_EXP_TIME, affectCell);
		//System.out.println(JSON.toJSONString(excelBook));
		System.out.println(step + "结束执行");
		logger.info("**********end excelId : "+excelId + " === step : " + step + "== reqPath : "+ reqPath);
		memcachedClient.set(excelId + "_ope", Constant.MEMCACHED_EXP_TIME, step);
	}
}
