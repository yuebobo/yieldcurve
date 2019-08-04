package com.mine;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

/**
 * 程序入口
 * GUI界面
 * 文件选择，获取基本路径，调用相应的处理方法
 * @author 
 *
 */
public class ButtonFrame extends JFrame{
	private static JFrame frame;
	
	public static void main(String[] args){
		EventQueue.invokeLater(() -> {
            frame = new ButtonFrame();
            frame.setTitle("数据处理");
            frame.setDefaultCloseOperation(EXIT_ON_CLOSE);
            frame.setVisible(true);
        });
	}

	private String basePath = null;
	private JFileChooser fic;
	private ButtonPanel buttonPanel;
	private static final int DEFAULT_WIDTH = 700;
	private static final int DEFAULT_HEIGHT = 400;

	public ButtonFrame(){
		setSize(DEFAULT_WIDTH,DEFAULT_HEIGHT);
		setLocationByPlatform(true);
		fic = new JFileChooser("请选择excel数据文件");
		buttonPanel = new ButtonPanel();
		buttonPanel.add(fic);
		add(buttonPanel);
		FileChoose choose = new FileChoose();
		fic.addActionListener(choose);
	}

	private class FileChoose implements ActionListener {

		@Override
		public void actionPerformed(ActionEvent e) {
			fic.disable();
			String filePath = fic.getSelectedFile().getName();
			String suffix = filePath.substring(filePath.lastIndexOf("."));
			if(".xlsx".equals(suffix)){
				basePath =fic.getSelectedFile().getPath();
				if(basePath != null && !"".equals(basePath)){
					try {
						Long st = System.currentTimeMillis();
						ExcelFile.dealFile(basePath);
						Long ed = System.currentTimeMillis();
						System.out.println("用时： "+(ed - st) +"毫秒");
						JOptionPane.showMessageDialog(buttonPanel, "已完成，请关闭窗口", "警告",JOptionPane.WARNING_MESSAGE);
					} catch (Exception ex) {
						JOptionPane.showMessageDialog(buttonPanel, "执行中发现异常", "警告",JOptionPane.WARNING_MESSAGE);
						ex.printStackTrace();
					}
				}
			}else {
				JOptionPane.showMessageDialog(buttonPanel, "该文件不是excel原始数据文件，请重新选择后缀为.xlsx的文件", "警告",JOptionPane.WARNING_MESSAGE);
			}
		}
	}

	class ButtonPanel extends JPanel{
		private static final int DEFAULT_WIDTH = 300;
		private static final int DEFAULT_HEIGHT = 200;

		@Override
		protected void paintComponent(Graphics g) {
			g.create();
			super.paintComponent(g);
		}

		@Override
		public Dimension getPreferredSize() {
			return new Dimension(DEFAULT_WIDTH, DEFAULT_HEIGHT);
		}
	}


}