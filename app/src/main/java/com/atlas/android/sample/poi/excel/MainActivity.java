package com.atlas.android.sample.poi.excel;

import android.Manifest;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;
import android.os.Bundle;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {
    private static final int REQUEST_CODE_PERMISSION = 100;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        // 动态申请读写外部存储文件权限
        if (ActivityCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE, Manifest.permission.READ_EXTERNAL_STORAGE},
                    REQUEST_CODE_PERMISSION);
            return;
        }

        exportExcelAsync();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (requestCode == REQUEST_CODE_PERMISSION) {
            if (resultCode == Activity.RESULT_OK) {
                exportExcelAsync();
            } else {
                Toast.makeText(this, "申请读写外部存储权限失败", Toast.LENGTH_LONG).show();
            }
        }
    }

    private void exportExcelAsync() {
        // 启动一个线程导出Excel表格
        new Thread(new Runnable() {
            @Override
            public void run() {
                boolean ret = false;

                try {
                    ret = exportExcel();
                } catch (Exception e) {
                    e.printStackTrace();
                }

                if (isFinishing()) {
                    return;
                }

                boolean finalRet = ret;
                runOnUiThread(new Runnable() {
                    @Override
                    public void run() {
                        Toast.makeText(MainActivity.this, finalRet ? "导出成功" : "导出失败", Toast.LENGTH_LONG).show();
                    }
                });
            }
        }).start();
    }

    private Person createPerson(String name, int age, String photoFileName) throws IOException {
        AssetManager assetManager = getAssets();
        InputStream is = assetManager.open(photoFileName);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        byte[] buffer = new byte[8 * 1024];
        int readlen;
        while ((readlen = is.read(buffer, 0, buffer.length)) != -1) {
            baos.write(buffer, 0, readlen);
        }
        is.close();

        return new Person(name, age, baos.toByteArray());
    }

    private boolean exportExcel() throws IOException {
        // 生成测试数据
        List<Person> personList = new ArrayList<>();
        personList.add(createPerson("张三", 27, "photo1.jpg"));
        personList.add(createPerson("李四", 35, "photo2.jpg"));

        // 设置有效数据的行数和列数
        int colNum = 3;   // ‘姓名’，‘年龄’，‘照片’三列
        int rowNum = personList.size();

        // 创建excel xlsx格式
        Workbook wb = new SXSSFWorkbook();
        // 创建工作表
        Sheet sheet = wb.createSheet();

        // 设置单元格显示宽度
        for (int i = 0; i < colNum; i++) {
            sheet.setColumnWidth(i, 20 * 256);  // 显示20个字符的宽度
        }

        // 设置单元格样式:居中显示
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        CreationHelper creationHelper = wb.getCreationHelper();

        for (int i = 0; i < rowNum; i++) {
            // 创建单元格
            Row row = sheet.createRow(i);
            // 设置单元格显示高度
            row.setHeightInPoints(128f);

            for (int j = 0; j < colNum; j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(cellStyle);

                if (j == 0) {
                    // 姓名
                    cell.setCellValue(personList.get(i).getName());
                } else if (j == 1) {
                    // 年龄
                    cell.setCellValue(personList.get(i).getAge());
                } else if (j == 2) {
                    // 照片
                    int picture = wb.addPicture(personList.get(i).getPhoto(), Workbook.PICTURE_TYPE_PNG);
                    Drawing drawingPatriarch = sheet.createDrawingPatriarch();
                    ClientAnchor anchor = creationHelper.createClientAnchor();
                    anchor.setCol1(j);
                    anchor.setRow1(i);
                    anchor.setCol2(j + 1);
                    anchor.setRow2(i + 1);
                    drawingPatriarch.createPicture(anchor, picture);
                }
            }
        }

        // 生成excel表格
        FileOutputStream fos = new FileOutputStream("/sdcard/person.xlsx");
        wb.write(fos);
        fos.flush();

        return true;
    }
}