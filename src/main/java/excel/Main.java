package excel;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        try {
            Workbook wb = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)wb.createSheet();

            // Excelシートに描画エリアを作成
            XSSFDrawing drawing = sheet.createDrawingPatriarch();

            // 線1の開始点と終了点を定義
            int x1 = 100;
            int y1 = 10;
            int x2 = 200;
            int y2 = 10;

            // 線1を描画
            drawLine(drawing, x1, y1, x2, y2);

            // 線2の開始点と終了点を定義（線1の終了点から始まる）
            int x3 = x2;
            int y3 = y2;
            int x4 = 300;
            int y4 = 300;

            // 線2を描画
            drawLine(drawing, x3, y3, x4, y4);

            // Excelファイルに書き込む
            FileOutputStream fileOut = new FileOutputStream("draw_line_example.xlsx");
            wb.write(fileOut);
            fileOut.close();

            // リソースを解放
            wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 線を描画するメソッド
    private static void drawLine(XSSFDrawing drawing, int x1, int y1, int x2, int y2) {
        XSSFClientAnchor anchor = new XSSFClientAnchor();
        // ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE == 左上座標と幅/高さの絶対値、セル参照なし
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        anchor.setDx1(Units.toEMU(x1)); // Start x-coordinate
        anchor.setDy1(Units.toEMU(y1)); // Start y-coordinate
        anchor.setDx2(Units.toEMU(x2)); // End x-coordinate
        anchor.setDy2(Units.toEMU(y2)); // End y-coordinate

        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(ShapeTypes.LINE);
        shape.setLineWidth(5.0); // 線の太さを設定
        shape.setLineStyleColor(0, 0, 255); // RGB値で線の色を設定 (青色)
    }

}
