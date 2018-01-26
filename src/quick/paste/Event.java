/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quick.paste;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 *
 * @author Bluemond
 * @throws AWTException
 */
public class Event implements ActionListener,
        Runnable {

    Form gui;
    Thread playing;
    QuickPaste x = new QuickPaste();
    int i;

    public Event(Form in) {
        gui = in;
    }

    @Override
    public void actionPerformed(ActionEvent event) {
        String command = event.getActionCommand();
        if (command.equals("读取")) {
            getFile();
        }
        if (command.equals("清除")) {
            clearFile();
        }
        if (command.equals("录入")) {
            gui.stop.setEnabled(true);
            try {
                pasteFile();
            } catch (Exception e) {
            }

        }
        if (command.equals("中止录入")) {
            playing = null;
            int j = i + 1;
            gui.got.setText("中止于第" + i + "/" + x.code.size() + "个条码");
            gui.paste.setEnabled(true);
            gui.stop.setEnabled(false);

        }
    }

    void getFile() {
        while (x.code.size() > 0) {
            x.code.remove(0);
        }
        x.testReadExcel(gui.address.getText());
        gui.got.setText("读取到" + x.code.size() + "个条码信息");
        try { // 防止文件建立或读取失败，用catch捕捉错误并打印，也可以throw  

            /* 写入Txt文件 */
            File writename = new File(".\\DataBox.txt"); // 相对路径，如果没有则要建立一个新的output.txt文件  
            writename.createNewFile(); // 创建新文件  
            BufferedWriter out = new BufferedWriter(new FileWriter(writename));
            out.write(gui.address.getText() + "\r\n"); // \r\n即为换行  
            out.write(gui.delay.getText());
            out.flush(); // 把缓存区内容压入文件  
            out.close(); // 最后记得关闭文件  

        } catch (IOException e) {
        }

    }

    void clearFile() {
        while (x.code.size() > 0) {
            x.code.remove(0);
        }
        gui.got.setText("已释放储存");
    }

    void pasteFile() throws Exception {
        playing = new Thread(this);
        playing.start();
        gui.paste.setEnabled(false);
    }

    @Override
    public void run() {
        Thread thisThread = Thread.currentThread();
        try {
            Robot robot = new Robot();
            KeyPress.keyPressWithAlt(robot, KeyEvent.VK_TAB);
            gui.got.setText("录入将在3秒后开始");
            robot.delay(3000);
            gui.got.setText("正在录入");
            int del = Integer.parseInt(gui.delay.getText());
            int j = x.code.size();
            for (i = 0; i < j; i++) {
                if (playing == thisThread) {

                    KeyPress.keyPressString(robot, x.code.get(i));
                    robot.delay(30);
                    KeyPress.keyPress(robot, KeyEvent.VK_ENTER); // 按下 enter 换行

                    robot.delay(del);
                } else {
                    break;
                }
            }
            if (i == x.code.size()) {
                gui.got.setText("已完成录入");
                gui.paste.setEnabled(true);
                gui.stop.setEnabled(false);
            }
        } catch (AWTException e) {
        }

    }

}
