package quick.paste;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.event.KeyEvent;

public class KeyPress {

    /**
     *
     * @param r
     *
     * @throws AWTException
     *
     */
    // shift+ 按键
    public static void keyPressWithShift(Robot r, int key) {
        r.keyPress(KeyEvent.VK_SHIFT);
        r.keyPress(key);
        r.keyRelease(key);
        r.keyRelease(KeyEvent.VK_SHIFT);
        r.delay(30);
    }

    // ctrl+ 按键
    public static void keyPressWithCtrl(Robot r, int key) {
        r.keyPress(KeyEvent.VK_CONTROL);
        r.keyPress(key);
        r.keyRelease(key);
        r.keyRelease(KeyEvent.VK_CONTROL);
    }

    // alt+ 按键
    public static void keyPressWithAlt(Robot r, int key) {
        r.keyPress(KeyEvent.VK_ALT);
        r.keyPress(key);
        r.keyRelease(key);
        r.keyRelease(KeyEvent.VK_ALT);
        r.delay(30);
    }

    //打印出字符串
    public static void keyPressString(Robot r, String str) {
        Clipboard clip = Toolkit.getDefaultToolkit().getSystemClipboard();//获取剪切板
        Transferable tText = new StringSelection(str);
        while (true) {
            try {
                clip.setContents(tText, null); //设置剪切板内容
                break;
            } catch (Exception e) {
                r.delay(100);
                System.out.print("出错已重做");
            }
        }
        keyPressWithCtrl(r, KeyEvent.VK_V);//粘贴
        r.delay(20);
    }

    //单个 按键
    public static void keyPress(Robot r, int key) {
        r.keyPress(key);
        r.keyRelease(key);
        r.delay(30);
    }

    public static boolean confirm1(Robot r) {
        Clipboard clip = Toolkit.getDefaultToolkit().getSystemClipboard();//获取剪切板
        Transferable tText = new StringSelection("");
        while (true) {
            try {
                clip.setContents(tText, null); //设置剪切板内容
                break;
            } catch (Exception e) {
            }
        }

        keyPressWithCtrl(r, KeyEvent.VK_A);
        r.delay(5);
        keyPressWithCtrl(r, KeyEvent.VK_C);
        if (clip.getContents(null) == null) {
            return true;

        }
        return false;

    }
     public static boolean confirm2(Robot r) {
        Clipboard clip = Toolkit.getDefaultToolkit().getSystemClipboard();//获取剪切板
        Transferable tText = new StringSelection("");
        while (true) {
            try {
                clip.setContents(tText, null); //设置剪切板内容
                break;
            } catch (Exception e) {
            }
        }

        keyPressWithCtrl(r, KeyEvent.VK_A);
        r.delay(5);
        keyPressWithCtrl(r, KeyEvent.VK_C);
        if (clip.getContents(null) != null) {
            return true;

        }
        return false;

    }
}
