package com.example.myapplication;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.content.pm.PackageManager;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;
import android.widget.EditText;
import android.widget.ImageView;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

public class MainActivity extends AppCompatActivity {

    private EditText editText;
    private ImageView imageView;

    private File filePath = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        ActivityCompat.requestPermissions(this,
                new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE,
                        Manifest.permission.READ_EXTERNAL_STORAGE},
                PackageManager.PERMISSION_GRANTED);

        editText = findViewById(R.id.editText);
        imageView = findViewById(R.id.imageView);

        filePath = new File(getExternalFilesDir(null), "ExcelFile.xlsx");

    }

    public void buttonInsertImage(View view) {

        try {
            String stringImageFilePath = Environment.getExternalStorageDirectory().getPath() +
                    "/Download/" + editText.getText().toString() + ".jpg";

            Bitmap bitmap = BitmapFactory.decodeFile(stringImageFilePath);
            imageView.setImageBitmap(bitmap);

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            bitmap.compress(Bitmap.CompressFormat.JPEG, 0, byteArrayOutputStream);
            byte[] bytesImage = byteArrayOutputStream.toByteArray();
            HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
            int intPictureIndex = hssfWorkbook.addPicture(bytesImage, Workbook.PICTURE_TYPE_JPEG);
            CreationHelper creationHelper = hssfWorkbook.getCreationHelper();

            ClientAnchor clientAnchor = creationHelper.createClientAnchor();
            clientAnchor.setCol1(6);
            clientAnchor.setRow1(19);
            clientAnchor.setCol2(7);
            clientAnchor.setRow2(20);

            HSSFSheet hssfSheet = hssfWorkbook.createSheet("local to excel");
            Drawing drawing = hssfSheet.createDrawingPatriarch();
            drawing.createPicture(clientAnchor, intPictureIndex);
            hssfSheet.createRow(1).createCell(1);
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            hssfWorkbook.write(fileOutputStream);

            if (fileOutputStream!=null){
                fileOutputStream.flush();
                fileOutputStream.close();
            }
            hssfWorkbook.close();

        }

        catch (Exception e) {

        }
    }
}