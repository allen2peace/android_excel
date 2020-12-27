package com.yjn.androidexcelreadwrite;

import android.app.ProgressDialog;
import android.content.Context;

public class ProgressDialogUtil {

	private static ProgressDialog pd;

	private static void show(Context context, String msg, boolean isCancel) {
		try {
			if (pd == null) {
				pd = new ProgressDialog(context);
			}
			pd.setCancelable(isCancel);
			pd.setMessage(msg + "");
			pd.show();
		} catch (Exception e) {
			destroy(context);
			e.printStackTrace();
		} catch (Error e) {
			destroy(context);
			e.printStackTrace();
		}
	}

	/**
	 * 获取pd对象
	 * 
	 * @param context
	 * @param msg
	 * @return
	 */
	public static void setMsg(Context context, String msg, boolean isCancel) {
		try {
			if (pd == null || !pd.isShowing()) {
				show(context, msg, isCancel);
			} else {
				pd.setMessage(msg + "");
			}
		} catch (Exception e) {
			destroy(context);
			e.printStackTrace();
		} catch (Error e) {
			destroy(context);
			e.printStackTrace();
		}
	}

	/**
	 * 消失对话框
	 * 
	 * @param context
	 */
	public static void dismiss(Context context) {
		try {
			if (pd != null && pd.isShowing()) {
				pd.dismiss();
			}
			pd = null;
		} catch (Exception e) {
			destroy(context);
			e.printStackTrace();
		}
	}

	/**
	 * 程序退出时要销毁此静态变量
	 * 
	 * @param context
	 */
	public static void destroy(Context context) {
		try {
			if (pd != null && pd.isShowing()) {
				pd.dismiss();
			}
			pd = null;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
