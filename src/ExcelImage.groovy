import java.awt.Color
import java.awt.image.BufferedImage

import javax.imageio.ImageIO

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook

// XSSFは2007以降の形式（xlsx）。2003以前の形式（xls）はHSSF
Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet()
workbook.setSheetName(0, "方眼紙")
sheet.setDisplayGridlines(false)

BufferedImage image = ImageIO.read(new File("../file/kappa100.png"));
def maxX = Math.min(200, image.getWidth())
def maxY = Math.min(200, image.getHeight())

// 無駄に見えてもセルの初期化をきちんとしてあげないといけない。
(0..<maxY).each {
	Row row = sheet.createRow(it);
	row.setHeightInPoints(2.5)
	(0..<maxX).each { row.createCell(it) }
}
(0..<maxX).each {
	sheet.setColumnWidth(it, 100)
}

println "作成開始"
for (int y: 0..<maxY) {
	Row row = sheet.getRow(y)
	for (int x: 0..<maxX) {
		int argb = image.getRGB(x, y)
		def rgb = [
			(argb >> 16) & 0xff,
			//red
			(argb >>  8) & 0xff,
			//green
			(argb      ) & 0xff  //blue
		]

		CellStyle style = workbook.createCellStyle()
		style.setFillPattern(CellStyle.SOLID_FOREGROUND)
		style.setFillForegroundColor(new XSSFColor(new Color(rgb[0], rgb[1], rgb[2])))
		Cell cell = row.getCell(x)
		cell.setCellStyle(style)
	}
	println "${y + 1} / ${maxY}"
}

new File("../file/image.xlsx").withOutputStream{workbook.write(it)}
