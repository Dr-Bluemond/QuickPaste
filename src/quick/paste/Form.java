/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quick.paste;

import java.awt.GridLayout;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.UIManager;

/**
 *
 * @author Bluemond
 */
public class Form extends JFrame {

    Event listener = new Event(this);
    JTextArea addressTextArea = new JTextArea("条码文件地址，默认为上次使用的地址，仅支持xlsx文件格式。");
    JTextField address = new JTextField();
    JButton get = new JButton("读取");
    JTextArea getTextArea = new JTextArea("输入xlsx文件地址后，请点击读取。");
    JTextField delay = new JTextField();
    JTextArea clearTextArea = new JTextArea("两次输入间的间隔，以毫秒为单位。");
    JButton paste = new JButton("录入");
    JTextArea pasteTextArea = new JTextArea("读取完成之后，打开扫书程序，再打开本界面，点击“录入”按钮，在接下来的三秒内点击扫书程序的输入框，等待即可。");
    JButton stop = new JButton("中止录入");
    JTextArea stopTextArea = new JTextArea("点击之后会中止录入并告知录入进程。");
    JTextField got = new JTextField();
    JTextArea gotTextArea = new JTextArea("信息面板:");

    public Form() {
        super("快捷录入器 v0.3 By 北京四中 朱绍铭");
        setLookAndFeel();
        setSize(510, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        GridLayout layout = new GridLayout(6, 2, 10, 10);
        setLayout(layout);

        try { // 防止文件建立或读取失败，用catch捕捉错误并打印，也可以throw  

            /* 读入TXT文件 */
            String pathname = ".\\DataBox.txt"; // 绝对路径或相对路径都可以，这里是绝对路径，写入文件时演示相对路径  
            File filename = new File(pathname); // 要读取以上路径的input。txt文件  
            InputStreamReader reader = new InputStreamReader(
                    new FileInputStream(filename)); // 建立一个输入流对象reader  
            BufferedReader br = new BufferedReader(reader); // 建立一个对象，它把文件内容转成计算机能读懂的语言  
            String line;
            line = br.readLine();
            address.setText(line);
            line = br.readLine();
            delay.setText(line);
        } catch (Exception e) {
            e.printStackTrace();
        }

        //add listeners
        get.addActionListener(listener);
        paste.addActionListener(listener);
        stop.addActionListener(listener);
        addressTextArea.setLineWrap(true);
        getTextArea.setLineWrap(true);
        clearTextArea.setLineWrap(true);
        pasteTextArea.setLineWrap(true);
        stopTextArea.setLineWrap(true);
        gotTextArea.setLineWrap(true);
        addressTextArea.setEditable(false);
        getTextArea.setEditable(false);
        clearTextArea.setEditable(false);
        pasteTextArea.setEditable(false);
        stopTextArea.setEditable(false);
        gotTextArea.setEditable(false);
        got.setEditable(false);
        stop.setEnabled(false);
        add(addressTextArea);
        add(address);
        add(getTextArea);
        add(get);
        add(clearTextArea);
        add(delay);
        add(pasteTextArea);
        add(paste);
        add(stopTextArea);
        add(stop);
        add(gotTextArea);
        add(got);

        setVisible(true);
    }

    private void setLookAndFeel() {
        try {
            UIManager.setLookAndFeel(
                    "com.sun.java.swing.plaf.nimbus.NimbusLookAndFeel"
            );
        } catch (Exception e) {

        }
    }
}
