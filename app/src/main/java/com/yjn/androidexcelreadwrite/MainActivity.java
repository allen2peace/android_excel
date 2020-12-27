package com.yjn.androidexcelreadwrite;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.OpenableColumns;
import android.text.TextUtils;
import android.util.Log;
import android.view.View;
import android.widget.TextView;
import android.widget.Toast;

import net.sourceforge.pinyin4j.PinyinHelper;
import net.sourceforge.pinyin4j.format.HanyuPinyinCaseType;
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat;
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType;
import net.sourceforge.pinyin4j.format.HanyuPinyinVCharType;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.recyclerview.widget.LinearLayoutManager;
import androidx.recyclerview.widget.RecyclerView;

public class MainActivity extends AppCompatActivity {
    public static final String TAG = MainActivity.class.getSimpleName();
    private Context mContext;
    private int FILE_SELECTOR_CODE = 10000;
    private int DIR_SELECTOR_CODE = 20000;
    private List<Map<Integer, Object>> readExcelList = new ArrayList<>();
    private RecyclerView recyclerView;
    private ExcelAdapter excelAdapter;
    private TextView tv_create_num;
    private TextView tv_record_num;

    int totalSize = 0;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        mContext = this;

        initViews();
        checkPermission(this, new IPermissionGrant() {
            @Override
            public void onGrant() {
                initCrashLogFolder(MainActivity.this);
            }
        });
    }

    private void initViews() {
        recyclerView = findViewById(R.id.excel_content_rv);
        excelAdapter = new ExcelAdapter(readExcelList);
        recyclerView.setLayoutManager(new LinearLayoutManager(this));
        recyclerView.setNestedScrollingEnabled(false);
        recyclerView.setAdapter(excelAdapter);
        tv_create_num = findViewById(R.id.tv_create_num);
        tv_record_num = findViewById(R.id.tv_record_num);
    }


    public void onClick(View view) {
        switch (view.getId()) {
            case R.id.import_excel_btn:
                openFileSelector();
                break;

            case R.id.export_excel_btn:
                if (readExcelList.size() > 0) {
                    openFolderSelector();
                } else {
                    Toast.makeText(mContext, "请先导入文件", Toast.LENGTH_SHORT).show();
                }
                break;
            default:
                break;
        }
    }

    /**
     * open local filer to select file
     */
    private void openFileSelector() {
        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
        intent.setType("application/*");
        startActivityForResult(intent, FILE_SELECTOR_CODE);
    }

    /**
     * open the local filer and select the folder
     */
    private void openFolderSelector() {
        Intent intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
        intent.setType("application/*");
        intent.putExtra(Intent.EXTRA_TITLE,
                System.currentTimeMillis() + ".xlsx");
        startActivityForResult(intent, DIR_SELECTOR_CODE);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == FILE_SELECTOR_CODE && resultCode == Activity.RESULT_OK) {
            Uri uri = data.getData();
            if (uri == null) return;
            Log.i(TAG, "onActivityResult: " + "filePath：" + uri.getPath());
            //select file and import
            importExcelDeal(uri);
        } else if (requestCode == DIR_SELECTOR_CODE && resultCode == Activity.RESULT_OK) {
            Uri uri = data.getData();
            if (uri == null) return;
            String uriPath = uri.getPath();
            Log.i(TAG, "onActivityResult: " + "filePath：" + uriPath);
            Toast.makeText(mContext, "导入中...", Toast.LENGTH_SHORT).show();
            //you can modify readExcelList, then write to excel.

            List<Map<Integer, Object>> tempList = new ArrayList<>();

            Map<Integer, Object> first = readExcelList.get(0);
//            Map<Integer, Object> second = readExcelList.get(1);
//            readExcelList.remove(0);
            readExcelList.remove(0);

            Collections.sort(readExcelList, new Comparator<Map<Integer, Object>>() {
                @Override
                public int compare(Map<Integer, Object> o1, Map<Integer, Object> o2) {
                    String name1 = getPingYin((String) o1.get(0));
                    String name2 = getPingYin((String) o2.get(0));
                    return name1.compareToIgnoreCase(name2);
                }
            });

            Log.d("tago", "onActivityResult: total size= " + readExcelList.size());

//            readExcelList.add(0, second);
            readExcelList.add(0, first);
            ArrayList<String> nameList = new ArrayList<>();
            ProgressDialogUtil.setMsg(this, "生成中,请等待...", false);
            new Thread(new Runnable() {
                @Override
                public void run() {
                    for (int j = 0; j < readExcelList.size(); j++) {
                        tempList.clear();

                        Log.d("tago", "开始检测: " + j);
                        Map<Integer, Object> item = readExcelList.get(j);
                        String name = (String) item.get(0);
                        if (nameList.contains(name)) {
                            Log.d("tago", "开始检测: " + j + ", return " + name);
                            continue;
                        }
                        nameList.add(name);
                        String s = initCrashLogFolder(MainActivity.this) + "/" + name + ".xlsx";

                        File file = new File(s);
                        try {
                            file.createNewFile();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                        Uri newUri = Uri.fromFile(file);

                        tempList.add(readExcelList.get(0));
//                tempList.add(readExcelList.get(1));

                        for (int k = j; k < readExcelList.size(); k++) {
                            Map<Integer, Object> innerItem = readExcelList.get(k);
                            String innerName = (String) innerItem.get(0);
                            if (TextUtils.equals(innerName, name)) {
                                tempList.add(innerItem);
                            } else {
                                break;
                            }
                        }

                        totalSize++;
                        Log.d("tago", "生成文件: " + name);
                        ExcelUtil.writeExcelNew(MainActivity.this, tempList, newUri);
                    }

                    runOnUiThread(() -> {
                        Toast.makeText(mContext, "导入成功!", Toast.LENGTH_SHORT).show();
                        tv_create_num.setText("导出文件数量: " + (totalSize - 1));
                        Log.d("tago", "onActivityResult: 生成文件总数: " + totalSize + ", " + readExcelList.size());
                        ProgressDialogUtil.dismiss(mContext);
                    });
                }
            }).start();
        }
    }

    public String getFileName(Uri uri) {
        String result = null;
        if (uri.getScheme().equals("content")) {
            Cursor cursor = getContentResolver().query(uri, null, null, null, null);
            try {
                if (cursor != null && cursor.moveToFirst()) {
                    result = cursor.getString(cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME));
                }
            } finally {
                cursor.close();
            }
        }
        if (result == null) {
            result = uri.getPath();
            int cut = result.lastIndexOf('/');
            if (cut != -1) {
                result = result.substring(cut + 1);
            }
        }
        return result;
    }

    public static String initCrashLogFolder(Context context) {
        String crashDir;
        File file;
        if (Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)) {
            file = new File(Environment.getExternalStorageDirectory() + "/Download/CeHuiWenjian");
        } else {
            file = new File(context.getCacheDir().getAbsolutePath() + "/CeHuiWenjian/");
        }
        if (!file.exists()) {
            file.mkdirs();
            if (!file.isDirectory()) {
                file.mkdirs();
            }
        }
        crashDir = file.getAbsolutePath();
        return crashDir;
    }

    private static IPermissionGrant permissionGrant;

    public static IPermissionGrant getPermissionGrant() {
        return permissionGrant;
    }

    public interface IPermissionGrant {
        void onGrant();
    }


    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        switch (requestCode) {
            case REQUEST_WRITE_EXTERNAL_STORAGE:
                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                    if (getPermissionGrant() != null) {
                        getPermissionGrant().onGrant();
                    }
                } else {
                    if (ActivityCompat.shouldShowRequestPermissionRationale(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
//                        PermissionUtils.showNeedPermissionDialog(this);
//                        ToolPermissionManager.showCusNeedPermissionDialog(this);
                    } else {
//                        goToSettingRequestPermission(this);
                    }
                }
                return;
        }
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
    }


    public static final int REQUEST_WRITE_EXTERNAL_STORAGE = 12;

    public static boolean checkPermission(Activity activity, IPermissionGrant pmGrant) {
        try {
            if (ContextCompat.checkSelfPermission(activity, Manifest.permission.READ_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED &&
                    ContextCompat.checkSelfPermission(activity, Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {
                return true;
            } else {
                // Permission is not granted, Should we show an explanation?
                if (ActivityCompat.shouldShowRequestPermissionRationale(activity, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {

                    // Show an explanation to the user *asynchronously* -- don't block
                    // this thread waiting for the user's response! After the user
                    // sees the explanation, try again to request the permission.
//                    showNeedPermissionDialog(activity);
                } else {
                    // No explanation needed; request the permission
                    ActivityCompat.requestPermissions(activity, new String[]{Manifest.permission.READ_EXTERNAL_STORAGE, Manifest.permission.WRITE_EXTERNAL_STORAGE}, REQUEST_WRITE_EXTERNAL_STORAGE);
                }

                permissionGrant = pmGrant;
                return false;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * 将字符串中的中文转化为拼音,其他字符不变
     */
    public static String getPingYin(String inputString) {
        HanyuPinyinOutputFormat format = new HanyuPinyinOutputFormat();
        format.setCaseType(HanyuPinyinCaseType.LOWERCASE);
        format.setToneType(HanyuPinyinToneType.WITHOUT_TONE);
        format.setVCharType(HanyuPinyinVCharType.WITH_V);

        char[] input = inputString.trim().toCharArray();// 把字符串转化成字符数组
        StringBuilder output = new StringBuilder();

        try {
            for (char c : input) {
                // \\u4E00是unicode编码，判断是不是中文
                if (Character.toString(c).matches("[\\u4E00-\\u9FA5]+")) {
                    // 将汉语拼音的全拼存到temp数组
                    String[] temp = PinyinHelper.toHanyuPinyinStringArray(
                            c, format);
                    // 取拼音的第一个读音
                    output.append(temp[0]);
                } else if (c > 'A' && c < 'Z') {
                    // 大写字母转化成小写字母
                    output.append(c);
                    output = new StringBuilder(output.toString().toLowerCase());
                }
                output.append(c);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return output.toString();
    }


    private void importExcelDeal(final Uri uri) {
        new Thread(() -> {
            Log.i(TAG, "doInBackground: Importing...");
            runOnUiThread(() -> Toast.makeText(mContext, "Importing...", Toast.LENGTH_SHORT).show());

//            List<Map<Integer, Object>> readExcelNew = ExcelUtil.readExcelNew(mContext, uri, uri.getPath());
            List<Map<Integer, Object>> readExcelNew = ExcelUtil.readExcelNew(mContext, uri, getFileName(uri));

            Log.i(TAG, "onActivityResult:readExcelNew " + ((readExcelNew != null) ? readExcelNew.size() : ""));

            if (readExcelNew != null && readExcelNew.size() > 0) {
                readExcelList.clear();
                readExcelList.addAll(readExcelNew);
                updateUI();

                Log.i(TAG, "run: successfully imported");
                runOnUiThread(() -> Toast.makeText(mContext, "导入成功", Toast.LENGTH_SHORT).show());
            } else {
                runOnUiThread(() -> Toast.makeText(mContext, "文件没有数据", Toast.LENGTH_SHORT).show());
            }
        }).start();
    }

    /**
     * refresh RecyclerView
     */
    private void updateUI() {
        runOnUiThread(() -> {

            tv_record_num.setText("发现记录条数:" + (readExcelList.size() - 1));
            if (readExcelList != null && readExcelList.size() > 0) {
                excelAdapter.notifyDataSetChanged();
            }
        });
    }

}
