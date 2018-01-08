package com.connor.HXom052;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class PPAPBean {
	private Integer index;// ���
	/** ��Ʒ�ͺ� */
	private String hx3_cpxh;
	/** ��Ʒ���� */
	private String hx3_cpmc;
	/** �������� */
	private String hx3_gfmc;
	/** Ҫ���ύʱ�� */
	private String hx3_yqtjsj;
	/** PPAP״̬ */
	private String hx3_ppapzt;
	/** ��׼״̬ */
	private String hx3_pzzt;
	/** ��ע */
	private String hx3_bz;
	/** ʵ���ύʱ�� */
	private String ratifyDate;// ������������Ϊ���ڱ��ϣ���Ҫ������
	/** bean���������� */
	public static final String[] Attr = { "hx3_cpxh", "hx3_cpmc", "hx3_gfmc", "hx3_yqtjsj", "hx3_ppapzt", "hx3_pzzt",
			"hx3_bz" };
	public static final String[] publicAttr = { "index", "hx3_cpxh", "hx3_cpmc", "hx3_gfmc", "hx3_yqtjsj", "ratifyDate",
			"hx3_ppapzt", "hx3_pzzt", "hx3_bz" };

	public PPAPBean() {
	}

	public Integer getIndex() {
		return index;
	}

	public void setIndex(Integer index) {
		this.index = index;
	}

	public String getHx3_cpxh() {
		return hx3_cpxh;
	}

	public void setHx3_cpxh(String hx3_cpxh) {
		this.hx3_cpxh = hx3_cpxh;
	}

	public String getHx3_cpmc() {
		return hx3_cpmc;
	}

	public void setHx3_cpmc(String hx3_cpmc) {
		this.hx3_cpmc = hx3_cpmc;
	}

	public String getHx3_gfmc() {
		return hx3_gfmc;
	}

	public void setHx3_gfmc(String hx3_gfmc) {
		this.hx3_gfmc = hx3_gfmc;
	}

	public String getHx3_yqtjsj() {
		return hx3_yqtjsj;
	}

	public void setHx3_yqtjsj(String hx3_yqtjsj) {
		this.hx3_yqtjsj = hx3_yqtjsj;
	}

	public String getHx3_ppapzt() {
		return hx3_ppapzt;
	}

	public void setHx3_ppapzt(String hx3_ppapzt) {
		this.hx3_ppapzt = hx3_ppapzt;
	}

	public String getHx3_pzzt() {
		return hx3_pzzt;
	}

	public void setHx3_pzzt(String hx3_pzzt) {
		this.hx3_pzzt = hx3_pzzt;
	}

	public String getHx3_bz() {
		return hx3_bz;
	}

	public void setHx3_bz(String hx3_bz) {
		this.hx3_bz = hx3_bz;
	}

	public String getRatifyDate() {
		return ratifyDate;
	}

	public void setRatifyDate(String ratifyDate) {
		this.ratifyDate = ratifyDate;
	}

	/**
	 * bean����������
	 * 
	 * @param attrName
	 * @param value
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public <T> void _setAttr(String attrName, T value) throws NoSuchMethodException, SecurityException,
			IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		String methodName = "set" + attrName.substring(0, 1).toUpperCase() + attrName.substring(1);
		Method method = this.getClass().getMethod(methodName, value.getClass());
		method.invoke(this, value);
	}

	/***
	 * bean����������
	 * 
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public void _printBean() throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException {
		Method method;
		StringBuilder sb = new StringBuilder();
		String methodName;
		for (String attr : Attr) {
			methodName = "get" + attr.substring(0, 1).toUpperCase() + attr.substring(1);
			method = this.getClass().getMethod(methodName);
			sb.append(attr + ":" + method.invoke(this) + "\t");
		}
		System.out.println(sb.toString());
	}

	@SuppressWarnings("unchecked")
	public <T> T _getBeanAttr(String attrName) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException {
		String methodName = "get" + attrName.substring(0, 1).toUpperCase() + attrName.substring(1);
		Method method = this.getClass().getMethod(methodName);
		return (T) method.invoke(this);

	}

	public static void main(String[] args) throws NoSuchMethodException, SecurityException, IllegalAccessException,
			IllegalArgumentException, InvocationTargetException {
		PPAPBean bean = new PPAPBean();
		bean._setAttr("index", 11);
		System.out.println(bean.getIndex());
	}
}
