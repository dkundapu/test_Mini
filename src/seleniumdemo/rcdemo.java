package seleniumdemo;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import java.util.concurrent.TimeUnit;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Color;
import java.awt.Font;
import java.awt.event.*;
import java.io.FileInputStream;
import java.io.IOException;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class rcdemo {
	static WebDriver driver;
	static FileInputStream fs;
	static Workbook wb;
	public static String first_party;
	public static String second_party;
	public static String amount;
	public static String description;
	public static String un;
	public static String ps;
	public static String ca;
	public static String file_path;
	public static JLabel l_error;
   public static void main(String[] args) throws InterruptedException {
	   driver = new ChromeDriver();
	   driver.manage().window().maximize();
	   TimeUnit.SECONDS.sleep(1);
	   driver.get("https://www.shcilestamp.com/eStampIndia/useradmin/UserAdminLoginServlet?rDoAction=LoadLoginPage");
	   try {
		   driver.findElement(By.name("userId"));
	   }catch(NullPointerException e){
		   System.out.println("No Login Page found.");
		   driver.quit();
	   }   
	   login();
       
       
        }
   public static void login() {
	   JFrame frame = new JFrame("Credentials");
	   frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	   JTextField uname, captcha, choosefile;
	   JLabel l_uname, l_captcha, l_password, l_choosefile;
	   l_uname = new JLabel("User Name:");
	   l_captcha = new JLabel("Captcha:");
	   l_password = new JLabel("Password:");
	   l_choosefile = new JLabel("File:");
	   l_error = new JLabel("");
	   l_error.setBounds(110, 200, 300, 50);
	   l_error.setForeground(Color.red);
	   l_uname.setBounds(10,10,100,30);
	   l_captcha.setBounds(10,110,100,30);
	   l_password.setBounds(10,60,100,30);
	   l_choosefile.setBounds(10, 160, 100, 30);
	   
	   uname=new JTextField(un);
	   uname.setBounds(110,10, 300,30); 
	   captcha = new JTextField(ca);
	   captcha.setBounds(110, 110, 300,30); 
	   JPasswordField pass = new JPasswordField(ps);  
	   pass.setBounds(110, 60, 300, 30);
	   choosefile = new JTextField(file_path);
	   choosefile.setBounds(110, 160, 300, 30);
	   choosefile.setEditable(false);
       frame.setSize(600,300);
       JButton b_login = new JButton("Continue");
       b_login.setBounds(420, 110, 90, 30);
       JButton b_choosefile = new JButton("Browse");
       b_choosefile.setBounds(420, 160, 90, 30);
       JFileChooser fc=new JFileChooser(file_path);
       fc.setBounds(110, 170, 300, 200);
       frame.add(b_login); // Adds Button to content pane of frame
       frame.add(uname);
       frame.add(captcha);
       frame.add(pass);
       frame.add(l_password);
       frame.add(l_uname);
       frame.add(l_captcha);
       frame.add(l_choosefile);
       frame.add(choosefile);
       frame.add(b_choosefile);
       frame.add(l_error);
       frame.setLayout(null);  
       frame.setVisible(true);
       
       b_login.addActionListener(new ActionListener() {
    	   public void actionPerformed(ActionEvent e) {
    		   l_error.setText("");
    		   un = uname.getText();
    		   ps = new String(pass.getPassword());
    		   ca = captcha.getText();
    		   //file_path = choosefile.getText();
    		   if (file_path == null){
    			   l_error.setText("Please select a file and hit 'continue'"); 
    		   }else {
    		   b_login.setEnabled(false);
    		   b_choosefile.setEnabled(false);
    		   l_error.setText("Note: Closing this window will stop the operation!!");
    		   driver.findElement(By.name("userId")).sendKeys(un);
    		   driver.findElement(By.name("userPwd")).sendKeys(ps);
    		   driver.findElement(By.name("searchjcaptcha")).sendKeys(ca);
    		   driver.findElement(By.name("searchjcaptcha")).sendKeys(Keys.ENTER);
    		   
    		   
    		   try {
    				readExcel();
    			} catch (BiffException f) {
    				// TODO Auto-generated catch block
    				f.printStackTrace();
    			} catch (IOException g) {
    				// TODO Auto-generated catch block
    				g.printStackTrace();
    			}
    		   }   
    	   }
       });
       
       b_choosefile.addActionListener(new ActionListener() {
    	   public void actionPerformed(ActionEvent e) {
    		   JFileChooser fc=new JFileChooser("Choose a file");
    		   FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel","xls");
    		   fc.setFileFilter(filter);
    		   fc.setFileSelectionMode(0);
    		   int i=fc.showOpenDialog(fc);
    		   if(i==JFileChooser.APPROVE_OPTION){
    		        file_path =fc.getSelectedFile().getPath();
    		        choosefile.setText(file_path);
    		   }

    		   
    	   }
       });
       
   }
   
   public static void readExcel() throws IOException, BiffException {
		String FilePath = file_path;
		
		try {
			fs = new FileInputStream(FilePath);
		} catch(IOException e){
			e.printStackTrace();
		}
		
		try {
			wb = Workbook.getWorkbook(fs);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// TO get the access to the sheet
		Sheet sh = wb.getSheet("Sheet1");

		// To get the number of rows present in sheet
		int totalNoOfRows = sh.getRows();

		// To get the number of columns present in sheet
		int totalNoOfCols = sh.getColumns();
		System.out.println("The file has "+ totalNoOfRows + " rows and "+ totalNoOfCols + " columns!!" );
		for (int row = 1; row < totalNoOfRows; row++) {
			first_party = sh.getCell(1, row).getContents().toString().toUpperCase();
			second_party = sh.getCell(2, row).getContents().toString().toUpperCase();
			description = sh.getCell(3, row).getContents().toString().toUpperCase();
			amount = sh.getCell(4, row).getContents().toString().toUpperCase();
			
			driver.get("https://www.shcilestamp.com/eStampIndia/submission/SubmissionServlet?rDoAction=LoadStampDuty");
			driver.findElement(By.name("RegSD")).sendKeys("B");
			driver.findElement(By.name("pNext")).sendKeys(Keys.SPACE);
			driver.findElement(By.id("TextField6Mand")).sendKeys(first_party);
			driver.findElement(By.id("TextField8Mand")).sendKeys(description);
			driver.findElement(By.id("TextField11Mand")).sendKeys(first_party);
			driver.findElement(By.id("TextField18Mand")).sendKeys(second_party);
			driver.findElement(By.id("TextField24Mand")).sendKeys(first_party);
			driver.findElement(By.id("TextField28Mand")).sendKeys(amount);
			driver.findElement(By.id("TextField28Mand")).sendKeys(Keys.TAB);
			
			try {
				TimeUnit.SECONDS.sleep(2);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			driver.switchTo().alert().accept();
			driver.findElement(By.name("pSave")).sendKeys(Keys.SPACE);
			driver.switchTo().alert().accept();
			
			
		}
		driver.get("https://www.shcilestamp.com/eStampIndia/useradmin/UserAdminMainServlet?rDoAction=Logout");
		driver.quit();
		l_error.setText("Done!! Please Close (X) this window now!!");
	}

}