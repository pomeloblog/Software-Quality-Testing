import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

public class Calculator extends JFrame{
    //用于判断是否重新开始
    private boolean start = true;
    private double result = 0;
    //存放加减乘除等于等
    private String command = "=";
    private JTextField jTextField;
    private JPanel jPanel = new JPanel();
    private JButton[] jButtons;
    //设置输出格式
    private DecimalFormat dof;

    //用构造方法进行必要的设置
    public Calculator() {

        this.setTitle("科学计算器");
        this.setSize(600, 400);
        this.setLocationRelativeTo(null);

        dof = new DecimalFormat("#.#####");//保留五位小数

        //添加文本域
        jTextField = new JTextField(30);
        jTextField.setText("");
        jTextField.setEditable(false);
        this.add(jTextField,"North");

        //添加按钮
        jPanel.setLayout(new GridLayout(6,7,3,3));
        String name[] = {
                "+/-","π","1/X","AC","/","*","DEL","X^2","X^3",
                "X^y","7","8","9","-","X!","√X","3^√X","4","5",
                "6","+","sin","cos","tan","1","2","3","%",
                "2进制","10进制","定积分","微分","0",".","=",
                "exp","log","ln","cot","time"," "," "
        };
        jButtons = new JButton[name.length];
        MyActionListener actionListener= new MyActionListener();

        //利用循环创建按钮对象并添加事件监听器
        for(int i = 0; i < name.length; i++) {

            jButtons[i] = new JButton(name[i]);
            jButtons[i].addActionListener(actionListener);

            //设置按钮背景颜色
            jButtons[i].setBackground(Color.lightGray);
            if(name[i].equals("="))
                jButtons[i].setBackground(Color.RED);
            else if((int)name[i].charAt(0)>=48 && (int)name[i].charAt(0)<=57
                    && name[i].length() == 1)
                jButtons[i].setBackground(Color.WHITE);
            // else if(name[i].equals("DEL"))
            //     jButtons[i].setBackground(Color.GRAY);

            jPanel.add(jButtons[i]);
        }

        this.add(jPanel);
        this.setDefaultCloseOperation(EXIT_ON_CLOSE);
        this.setVisible(true);

    }

    //用内部类实现事件监听器
    class MyActionListener implements ActionListener {
        //按钮被单击
        public void actionPerformed(ActionEvent e) {

            String input = e.getActionCommand();
            //开始
            if(start) {
                if((int)input.charAt(0)>=48 && (int)input.charAt(0)<=57
                        && input.length() == 1 ) {
                    jTextField.setText(""+input);
                }
                if(input.equals("+/-")) {
                    jTextField.setText("-");
                }
                if(input.equals("π")) {
                    jTextField.setText(""+Math.PI);
                }
                start = false;
                if(input.equals("AC"))
                    jTextField.setText("");
            }
            //0~9数字等非运算符
            else if((int)input.charAt(0)>=48 && (int)input.charAt(0)<=57
                    && input.length() == 1 || input.equals(".")){
                jTextField.setText(jTextField.getText()+input);
            }
            //实现所有 数组+运算符 的运算

            //实现清零键
            else if(input.equals("AC"))
                jTextField.setText("");
                //实现退格键
            else if(input.equals("DEL")) {
                if(jTextField.getText().length() > 0){
                    jTextField.setText(jTextField.getText().substring(0,jTextField.getText().length()-1));
                }
            }
            //实现正弦三角函数
            else if(input.equals("sin")) {
                result = Math.sin(Double.parseDouble(jTextField.getText()));
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                jTextField.setText(dof.format(result));
                start = true;
            }
            //实现余弦三角函数
            else if(input.equals("cos")) {
                result = Math.cos(Double.parseDouble(jTextField.getText()));
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现余切三角函数
            else if(input.equals("cot")) {
                result = 1.0/Math.tan(Double.parseDouble(jTextField.getText()));
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现正切三角函数
            else if(input.equals("tan")) {
                result = Math.tan(Double.parseDouble(jTextField.getText()));
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现十进制到二进制的转化
            else if(input.equals("2进制")) {
                String result2 = Integer.toBinaryString(Integer.parseInt(jTextField.getText()));
                jTextField.setText(""+getPrettyNumber(result2));
                start = true;
            }
            //实现二进制到十进制的转化
            else if(input.equals("10进制")) {
                try {
                    String result2 = Integer.valueOf(jTextField.getText(),2).toString();
                    jTextField.setText(""+getPrettyNumber(result2));
                }catch(NumberFormatException exception) {
                    JOptionPane.showMessageDialog(null, "对不起，数字错误，请重新输入！", "Error!", JOptionPane.ERROR_MESSAGE);
                    throw new NumberFormatException("数字格式错误");
                }finally {
                    start = true;
                }

            }
            //实现1/x
            else if(input.equals("1/X")) {
                result = 1 / Double.parseDouble(jTextField.getText());
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现平方计算
            else if(input.equals("X^2")) {
                result = Math.pow(Double.parseDouble(jTextField.getText()), 2);
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现立方计算
            else if(input.equals("X^3")) {
                result = Math.pow(Double.parseDouble(jTextField.getText()), 3);
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现阶乘
            else if(input.equals("X!")) {
                if(Double.parseDouble(jTextField.getText()) < 0) {
                    JOptionPane.showMessageDialog(null, "对不起，阶乘计算不能为负数", "Error!", JOptionPane.ERROR_MESSAGE);
                    jTextField.setText("对不起，阶乘计算不能为负数");
                    start = true;
                    throw new IllegalArgumentException("阶乘计算出现负数");
                }else {
                    int sum;
                    sum = factorial(Integer.parseInt(jTextField.getText()));
                    jTextField.setText(Integer.toString(sum));
                    start = true;
                }

            }
            //实现百分号计算
            else if(input.equals("%")) {
                result = Double.parseDouble(jTextField.getText())/ 100.0;
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现开平方根
            else if(input.equals("√X")) {
                if(Double.parseDouble(jTextField.getText()) < 0){
                    JOptionPane.showMessageDialog(null, "负数不能开根", "Error!", JOptionPane.ERROR_MESSAGE);
                    jTextField.setText("对不起，开方计算不能为负数");
                    start = true;
                    throw new IllegalArgumentException("开方计算出现负数");
                }else{
                    result = Math.sqrt(Double.parseDouble(jTextField.getText()));
                    jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                    start = true;
                }

            }
            //实现开立方根
            else if(input.equals("3^√X")) {
                result = Math.pow(Double.parseDouble(jTextField.getText()),1.0/3);
                jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                start = true;
            }
            //实现获取当前时间
            else if(input.equals("time")) {
                SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                jTextField.setText(df.format(System.currentTimeMillis()));
                start = true;
            }
            else if(input.equals("微分")){
                String title = JOptionPane.showInputDialog(null, "请输入一元二次方程的三个参数和微分点，以空格分开：\n", "title", JOptionPane.PLAIN_MESSAGE);
                String[] strs = title.split(" ");
                result = derivative(Double.parseDouble(jTextField.getText()),0.0000001, strs);
                // jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                jTextField.setText(dof.format(result));
                start = true;
            }
            else if(input.equals("exp")){
                result = Math.exp(Double.parseDouble(jTextField.getText()));
                jTextField.setText(dof.format(result));
                start = true;
            }
            else if(input.equals("ln")){
                result = Math.log(Double.parseDouble(jTextField.getText()));
                jTextField.setText(dof.format(result));
                start = true;
            }
            else if(input.equals("定积分")){
                String title = JOptionPane.showInputDialog(null, "请输入一元二次方程的三个参数和积分区间，以空格分开：\n", "title", JOptionPane.PLAIN_MESSAGE);
                String[] str = title.split(" ");
                // result = defintegration(result, Double.parseDouble(jTextField.getText()), 10000);
                result = defintegration(Double.parseDouble(str[3]), Double.parseDouble(str[4]), 10000, str);
                jTextField.setText(dof.format(result));
                start = true;
            }

            //实现加减乘除等 数字+运算符+数字 形式的运算
            else {
                if(!start) {

                    if(command.equals("+"))
                        result += Double.parseDouble(jTextField.getText());
                    else if(command.equals("-"))
                        result -= Double.parseDouble(jTextField.getText());
                    else if(command.equals("*"))
                        result *= Double.parseDouble(jTextField.getText());
                    else if(command.equals("/")) {
                        if(Double.parseDouble(jTextField.getText()) != 0) {
                            result /= Double.parseDouble(jTextField.getText());
                        }else {
                            jTextField.setText(""+"对不起，除数不能为零");
                            JOptionPane.showMessageDialog(null, "对不起，除数不能为零", "Error!", JOptionPane.ERROR_MESSAGE);
                            command = "=";
                            start = true;
                            throw new ArithmeticException("除数为零");
                        }

                    }
                    else if(command.equals("="))
                        result = Double.parseDouble(jTextField.getText());
                    else if(command.equals("X^y"))
                        result = Math.pow(result, Double.parseDouble(jTextField.getText()));
                    else if(command.equals("log"))  //先输入底数再输入真数
                        result = Math.log( Double.parseDouble(jTextField.getText()))/Math.log(result);
                    // jTextField.setText(""+getPrettyNumber(Double.toString(result)));
                    jTextField.setText(dof.format(result));
                    command = input;
                    start = true;
                }
            }
        }
    }
    //去掉小数点后没用的0
    public static String getPrettyNumber(String number) {
        return BigDecimal.valueOf(Double.parseDouble(number))
                .stripTrailingZeros().toPlainString();
    }
    //用循环计算阶乘
    public static int factorial(int num) {
        int sum = 1;
        for(int i = 1;i <= num; i++){
            sum *= i;
        }
        return sum;
    }
    //面积法计算定积分
    public static double defintegration(double a, double b, int accuracy, String[] str){
        double sum = 0;
        // int accuracy = ;
        double step = (b - a)/accuracy;
        for(int i = 1;i <= accuracy; i++){
            double x = a + step*i;
            sum = sum + f(x,str) * step;
        }
        return sum;
    }
    //定义函数
    public static double f(double x, String[] str){
        return Double.parseDouble(str[0])*x*x + Double.parseDouble(str[1])*x + Double.parseDouble(str[2]);
    }
    //定义法求微分
    public static double derivative(double a,double accuracy, String[] str){
        return (f(a + accuracy, str)-f(a - accuracy, str))/(2*accuracy);
    }

    public static void main(String[] args) {
        Calculator Calculator = new Calculator();

    }

}

