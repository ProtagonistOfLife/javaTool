package com.util.excel;


import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextField;

public class CompareWind extends JFrame {

	private static final long serialVersionUID = 1L;
	private JTextField jf1,jf2;
	private JButton jb;
	private JPanel jp;
	private JTextArea ja;
	private JDialog jd;
	private JLabel jl;
	
	public CompareWind() {
		this.setLayout(new BorderLayout());
		init();
		this.setBounds(200, 100, 300, 200);
		this.setVisible(true);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}
	
	public void init(){
		jp = new JPanel(new GridLayout(2, 1));
		jf1 = new JTextField(100);
		jf2 = new JTextField(100);
		ja = new JTextArea(4, 100);
		jb = new JButton("开始对比");
		jd = new JDialog(this, "结果", true);
		jl = new JLabel();
		
		ja.setEditable(false);
		
		jp.add(jf1);
		jp.add(jf2);
		this.add(jp, BorderLayout.NORTH);
		this.add(ja, BorderLayout.CENTER);
		this.add(jb, BorderLayout.SOUTH);
		
		jb.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				ExcelCompare ec = new ExcelCompare();
				
				String path1 = jf1.getText();
				String path2 = jf2.getText();
				jd.setBounds(250, 125, 100, 50);
				jd.setLayout(new BorderLayout());
				if(ec.fileCompare(path1, path2) == 0){
					jl.setText("相同");
					jd.add(jl);
					ja.setText("相同");
					jd.setVisible(true);
				}else if(ec.fileCompare(path1, path2) == 1){
					jl.setText("不同");
					jd.setVisible(true);
					StringBuilder sb = new StringBuilder("不同");
					Map<String, String> map = null;
					sb.append("\r\n");
					if((map=ec.getMap()) != null){
						for(Map.Entry<String, String> en : map.entrySet()){
							sb.append(en.getKey() + "->" + en.getValue());
							sb.append("\r\n");
						}
					}
					ja.setText(sb.toString());
				}else{
					jl.setText("比较异常");
					jd.setVisible(true);
					ja.setText("比较异常");
				}
			}
		});
	}
	
	public static void main(String[] args) {
		new CompareWind();
	}
}
