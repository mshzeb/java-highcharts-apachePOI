import java.io.*;

import java.util.Iterator;

import javax.swing.ImageIcon;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.util.IOUtils;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;


public class CreatePieChartExample {

	public static void main(String[] args) throws Exception {

		FileInputStream chart_file_input = new FileInputStream(new File("D:\\Java Projects Juno\\ApachePOI\\src\\my_chart.xls"));
		HSSFWorkbook my_workbook = new HSSFWorkbook(chart_file_input);
		HSSFSheet my_sheet = my_workbook.getSheetAt(0);
		
		DefaultPieDataset my_pie_chart_data = new DefaultPieDataset();
		
		Iterator<Row> rowIterator = my_sheet.iterator();
		String chart_label = "c";
		Number chart_data = 0;
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch(cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
							chart_data = cell.getNumericCellValue();
							break;
					case Cell.CELL_TYPE_STRING:
							chart_label = cell.getStringCellValue();
							break;
					}
				}
				my_pie_chart_data.setValue(chart_label, chart_data);
		}
		JFreeChart myPieChart = ChartFactory.createPieChart("Excel Pie Chart Java Example", my_pie_chart_data, true, true, false);
		int width = 640, height = 480;
		float quality = 1;
		ByteArrayOutputStream chart_out = new ByteArrayOutputStream();
		ChartUtilities.writeChartAsJPEG(chart_out, quality, myPieChart, width, height);
		
		InputStream feed_chart_to_excel = new ByteArrayInputStream(chart_out.toByteArray());
		byte[] bytes = IOUtils.toByteArray(feed_chart_to_excel);
		int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		feed_chart_to_excel.close();
		chart_out.close();
		
		HSSFPatriarch drawing = my_sheet.createDrawingPatriarch();
		ClientAnchor my_anchor = new HSSFClientAnchor();
		
		my_anchor.setCol1(4);
		my_anchor.setRow1(5);
		
		HSSFPicture my_picture = drawing.createPicture(my_anchor, my_picture_id);
		my_picture.resize();
		
		FileOutputStream out = new FileOutputStream(new File("D:\\Java Projects Juno\\ApachePOI\\src\\my_chart.xls"));
		my_workbook.write(out);
		out.close();
		
		ImageIcon icon = new ImageIcon("images/check-32x32.png");
		JOptionPane.showMessageDialog(null, "Chart successfully generated programmatically.", "Java Apache POI Info", JOptionPane.INFORMATION_MESSAGE, icon);

	}

}